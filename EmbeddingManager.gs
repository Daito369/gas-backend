/**
 * RAG 2.0 エンベディング管理モジュール
 * 埋め込みベクトルの生成と管理を担当します
 */
class EmbeddingManager {

  /**
   * チャンクのエンベディングを生成します
   * @param {string} chunkId チャンクID
   * @return {Object} 生成結果
   */
  static async generateEmbeddingForChunk(chunkId) {
    try {
      // チャンクを取得
      const chunk = SheetStorage.getChunkById(chunkId);
      if (!chunk) {
        throw new Error(`チャンク ${chunkId} が見つかりません`);
      }

      // チャンクの内容を取得
      const content = chunk.content;
      if (!content || content.trim().length === 0) {
        throw new Error(`チャンク ${chunkId} の内容が空です`);
      }

      // エンベディングを生成
      const result = await GeminiIntegration.generateEmbedding(content);
      
      if (!result.success || !result.embedding) {
        throw new Error(`エンベディング生成に失敗しました: ${result.error || 'Unknown error'}`);
      }

      // エンベディングを保存
      const saveResult = SheetStorage.saveEmbedding({
        chunk_id: chunkId,
        document_id: chunk.document_id,
        vector: result.embedding,
        category: chunk.category || 'general',
        model_version: result.model
      });

      if (!saveResult.success) {
        throw new Error(`エンベディングの保存に失敗しました: ${saveResult.error || 'Unknown error'}`);
      }

      return {
        success: true,
        chunk_id: chunkId,
        document_id: chunk.document_id,
        embedding_size: result.embedding.length,
        model: result.model
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.generateEmbeddingForChunk',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_id: chunkId },
        retry: true
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * ドキュメントの全チャンクのエンベディングを生成します
   * @param {string} documentId ドキュメントID
   * @return {Object} 生成結果
   */
  static async generateEmbeddingsForDocument(documentId) {
    try {
      // ドキュメントのチャンクを取得
      const chunks = SheetStorage.getChunksByDocumentId(documentId);
      if (!chunks || chunks.length === 0) {
        throw new Error(`ドキュメント ${documentId} のチャンクが見つかりません`);
      }

      const results = {
        success: true,
        document_id: documentId,
        total_chunks: chunks.length,
        processed_chunks: 0,
        success_chunks: 0,
        failed_chunks: 0,
        errors: []
      };

      // レート制限を考慮した処理
      for (const chunk of chunks) {
        try {
          // エンベディング生成
          const result = await this.generateEmbeddingForChunk(chunk.id);
          results.processed_chunks++;
          
          if (result.success) {
            results.success_chunks++;
          } else {
            results.failed_chunks++;
            results.errors.push({
              chunk_id: chunk.id,
              error: result.error
            });
          }
          
          // API制限を考慮して少し待機
          if (results.processed_chunks < chunks.length) {
            Utilities.sleep(500);
          }
        } catch (error) {
          results.failed_chunks++;
          results.errors.push({
            chunk_id: chunk.id,
            error: error.message
          });
        }
      }

      // 全体の成功判定を更新
      results.success = results.failed_chunks === 0;

      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.generateEmbeddingsForDocument',
        error: error,
        severity: 'HIGH',
        context: { document_id: documentId }
      });

      return {
        success: false,
        document_id: documentId,
        error: error.message
      };
    }
  }

  /**
   * カテゴリの全ドキュメントのエンベディングを生成します
   * @param {string} category カテゴリ
   * @param {Object} options オプション
   * @return {Object} 生成結果
   */
  static async generateEmbeddingsForCategory(category, options = {}) {
    try {
      // カテゴリのドキュメントリストを取得
      const documents = this.getDocumentsByCategory(category);
      if (!documents || documents.length === 0) {
        throw new Error(`カテゴリ ${category} のドキュメントが見つかりません`);
      }

      // デフォルトオプション
      const defaultOptions = {
        maxDocuments: 0, // 0 = 無制限
        skipExisting: true,
        batchSize: 5,
        delayBetweenBatches: 5000 // 5秒
      };

      // オプションをマージ
      const mergedOptions = { ...defaultOptions, ...options };

      const results = {
        success: true,
        category: category,
        total_documents: documents.length,
        processed_documents: 0,
        success_documents: 0,
        failed_documents: 0,
        skipped_documents: 0,
        errors: []
      };

      // ドキュメント処理対象を制限
      let documentsToProcess = documents;
      if (mergedOptions.maxDocuments > 0 && documents.length > mergedOptions.maxDocuments) {
        documentsToProcess = documents.slice(0, mergedOptions.maxDocuments);
        results.total_documents = mergedOptions.maxDocuments;
      }

      // バッチ処理
      for (let i = 0; i < documentsToProcess.length; i += mergedOptions.batchSize) {
        const batch = documentsToProcess.slice(i, i + mergedOptions.batchSize);
        
        for (const document of batch) {
          try {
            // 既存のエンベディングをスキップ
            if (mergedOptions.skipExisting) {
              const hasEmbeddings = await this.documentHasEmbeddings(document.id);
              if (hasEmbeddings) {
                results.skipped_documents++;
                results.processed_documents++;
                continue;
              }
            }
            
            // エンベディング生成
            const result = await this.generateEmbeddingsForDocument(document.id);
            results.processed_documents++;
            
            if (result.success) {
              results.success_documents++;
            } else {
              results.failed_documents++;
              results.errors.push({
                document_id: document.id,
                error: result.error
              });
            }
          } catch (error) {
            results.failed_documents++;
            results.processed_documents++;
            results.errors.push({
              document_id: document.id,
              error: error.message
            });
          }
        }
        
        // バッチ間の遅延
        if (i + mergedOptions.batchSize < documentsToProcess.length) {
          Utilities.sleep(mergedOptions.delayBetweenBatches);
        }
      }

      // 全体の成功判定を更新
      results.success = results.failed_documents === 0;

      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.generateEmbeddingsForCategory',
        error: error,
        severity: 'HIGH',
        context: { category: category }
      });

      return {
        success: false,
        category: category,
        error: error.message
      };
    }
  }

  /**
   * すべてのカテゴリのエンベディングを生成します
   * @param {Object} options オプション
   * @return {Object} 生成結果
   */
  static async generateAllEmbeddings(options = {}) {
    try {
      // すべてのカテゴリを取得
      const categories = this.getAllCategories();
      
      const results = {
        success: true,
        total_categories: categories.length,
        processed_categories: 0,
        success_categories: 0,
        failed_categories: 0,
        category_results: {},
        errors: []
      };

      // 各カテゴリを処理
      for (const category of categories) {
        try {
          const categoryResult = await this.generateEmbeddingsForCategory(category, options);
          results.processed_categories++;
          
          if (categoryResult.success) {
            results.success_categories++;
          } else {
            results.failed_categories++;
            results.errors.push({
              category: category,
              error: categoryResult.error
            });
          }
          
          // カテゴリの結果を保存
          results.category_results[category] = {
            total_documents: categoryResult.total_documents,
            processed_documents: categoryResult.processed_documents,
            success_documents: categoryResult.success_documents,
            failed_documents: categoryResult.failed_documents,
            skipped_documents: categoryResult.skipped_documents
          };
        } catch (error) {
          results.failed_categories++;
          results.processed_categories++;
          results.errors.push({
            category: category,
            error: error.message
          });
        }
      }

      // 全体の成功判定を更新
      results.success = results.failed_categories === 0;

      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.generateAllEmbeddings',
        error: error,
        severity: 'HIGH',
        context: { options: JSON.stringify(options) }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * クエリテキストのエンベディングを生成します
   * @param {string} queryText クエリテキスト
   * @return {Object} 生成結果
   */
  static async generateQueryEmbedding(queryText) {
    try {
      if (!queryText || queryText.trim().length === 0) {
        throw new Error('クエリテキストが空です');
      }

      // クエリの言語を検出
      const language = Utilities.detectLanguage(queryText);

      // 言語に応じた前処理
      let processedQuery = queryText;
      if (language === 'ja') {
        processedQuery = Utilities.preprocessJapaneseText(queryText);
      } else {
        processedQuery = Utilities.preprocessEnglishText(queryText);
      }

      // エンベディング生成
      const result = await GeminiIntegration.generateEmbedding(processedQuery);
      
      if (!result.success || !result.embedding) {
        throw new Error(`クエリエンベディング生成に失敗しました: ${result.error || 'Unknown error'}`);
      }

      return {
        success: true,
        embedding: result.embedding,
        model: result.model,
        language: language,
        processed_query: processedQuery
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.generateQueryEmbedding',
        error: error,
        severity: 'MEDIUM',
        context: { query_length: queryText ? queryText.length : 0 },
        retry: true
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * ドキュメントが既にエンベディングを持っているか確認します
   * @param {string} documentId ドキュメントID
   * @return {boolean} エンベディングが存在するかどうか
   */
  static async documentHasEmbeddings(documentId) {
    try {
      // ドキュメントのチャンクを取得
      const chunks = SheetStorage.getChunksByDocumentId(documentId);
      if (!chunks || chunks.length === 0) {
        return false;
      }

      // 最初のチャンクだけチェック（パフォーマンスのため）
      const firstChunk = chunks[0];
      const embedding = SheetStorage.getEmbeddingByChunkId(firstChunk.id);
      
      return !!embedding;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.documentHasEmbeddings',
        error: error,
        severity: 'LOW',
        context: { document_id: documentId }
      });
      
      return false;
    }
  }

  /**
   * カテゴリのドキュメントリストを取得します
   * @param {string} category カテゴリ
   * @return {Array<Object>} ドキュメントリスト
   */
  static getDocumentsByCategory(category) {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const documentsSheet = ss.getSheetByName('Documents');
      
      if (!documentsSheet) {
        throw new Error('Documents シートが見つかりません');
      }
      
      // データを取得
      const data = documentsSheet.getDataRange().getValues();
      if (data.length <= 1) {
        return [];
      }
      
      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const idIndex = headers.indexOf('id');
      const titleIndex = headers.indexOf('title');
      const categoryIndex = headers.indexOf('category');
      
      if (idIndex === -1 || categoryIndex === -1) {
        throw new Error('Documents シートの形式が不正です');
      }
      
      // カテゴリに一致するドキュメントを検索
      const documents = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[categoryIndex] === category) {
          const document = {
            id: row[idIndex],
            title: titleIndex !== -1 ? row[titleIndex] : '',
            category: category
          };
          documents.push(document);
        }
      }
      
      return documents;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.getDocumentsByCategory',
        error: error,
        severity: 'MEDIUM',
        context: { category: category }
      });
      
      return [];
    }
  }

  /**
   * すべてのカテゴリを取得します
   * @return {Array<string>} カテゴリリスト
   */
  static getAllCategories() {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const documentsSheet = ss.getSheetByName('Documents');
      
      if (!documentsSheet) {
        throw new Error('Documents シートが見つかりません');
      }
      
      // データを取得
      const data = documentsSheet.getDataRange().getValues();
      if (data.length <= 1) {
        return [];
      }
      
      // ヘッダー行を取得してカテゴリのインデックスをマッピング
      const headers = data[0];
      const categoryIndex = headers.indexOf('category');
      
      if (categoryIndex === -1) {
        throw new Error('Documents シートの形式が不正です');
      }
      
      // 一意のカテゴリを収集
      const categorySet = new Set();
      for (let i = 1; i < data.length; i++) {
        const category = data[i][categoryIndex];
        if (category) {
          categorySet.add(category);
        }
      }
      
      return Array.from(categorySet);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.getAllCategories',
        error: error,
        severity: 'MEDIUM'
      });
      
      // デフォルトカテゴリリストを返す
      return [
        'Help_Pages',
        'Search',
        'Mobile',
        'Shopping',
        'Display',
        'Video',
        'M&A',
        'Billing',
        'Policy',
        'general'
      ];
    }
  }

  /**
   * 操作を再試行します
   * @param {Object} options 再試行オプション
   * @return {boolean} 再試行が成功したかどうか
   */
  static async retryOperation(options) {
    try {
      const { operation, params } = options.context;
      
      // 操作に基づいて適切なメソッドを呼び出す
      switch (operation) {
        case 'generateEmbeddingForChunk':
          const chunkResult = await this.generateEmbeddingForChunk(params.chunkId);
          return chunkResult.success;
          
        case 'generateQueryEmbedding':
          const queryResult = await this.generateQueryEmbedding(params.queryText);
          return queryResult.success;
          
        default:
          console.warn(`Unknown operation for retry: ${operation}`);
          return false;
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.retryOperation',
        error: error,
        severity: 'HIGH',
        context: options.context
      });
      
      return false;
    }
  }

  /**
   * Colabスクリプトを使用してバッチエンベディングを実行するための設定を生成します
   * @param {string} category カテゴリ（省略時は全カテゴリ）
   * @param {Object} options オプション
   * @return {Object} Colab設定
   */
  static generateColabConfig(category = null, options = {}) {
    try {
      // デフォルトオプション
      const defaultOptions = {
        maxDocuments: 0, // 0 = 無制限
        skipExisting: true,
        batchSize: 50,
        outputPath: 'embeddings_output.json'
      };

      // オプションをマージ
      const mergedOptions = { ...defaultOptions, ...options };

      // Colab用の設定を作成
      const config = {
        database_id: Config.getDatabaseId(),
        api_key: "YOUR_API_KEY_HERE", // Colabでは手動で設定
        embedding_model: Config.getSystemConfig().embedding_model,
        categories: category ? [category] : this.getAllCategories(),
        options: {
          max_documents: mergedOptions.maxDocuments,
          skip_existing: mergedOptions.skipExisting,
          batch_size: mergedOptions.batchSize
        },
        output_path: mergedOptions.outputPath
      };

      return {
        success: true,
        config: config,
        instructions: [
          "1. Colabノートブックで以下の設定を使用してください",
          "2. YOUR_API_KEY_HERE を実際のGemini API キーに置き換えてください",
          "3. 実行後、出力ファイルをGoogleドライブにダウンロードして、インポートスクリプトを実行してください"
        ]
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.generateColabConfig',
        error: error,
        severity: 'MEDIUM',
        context: { category: category }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Colabで生成されたエンベディングをインポートします
   * @param {string} fileId エンベディングファイルのID
   * @return {Object} インポート結果
   */
  static importEmbeddingsFromColab(fileId) {
    try {
      // ファイルを取得
      const file = DriveApp.getFileById(fileId);
      if (!file) {
        throw new Error(`ファイル ${fileId} が見つかりません`);
      }

      // ファイル内容を取得
      const fileContent = file.getBlob().getDataAsString();
      const embeddings = JSON.parse(fileContent);

      // 結果初期化
      const results = {
        success: true,
        total_embeddings: embeddings.length,
        imported_count: 0,
        failed_count: 0,
        errors: []
      };

      // エンベディングをインポート
      for (const item of embeddings) {
        try {
          // エンベディングを保存
          const saveResult = SheetStorage.saveEmbedding({
            chunk_id: item.chunk_id,
            document_id: item.document_id,
            vector: item.embedding,
            category: item.category || 'general',
            model_version: item.model_version || Config.getSystemConfig().embedding_model
          });

          if (saveResult.success) {
            results.imported_count++;
          } else {
            results.failed_count++;
            results.errors.push({
              chunk_id: item.chunk_id,
              error: saveResult.error
            });
          }
        } catch (error) {
          results.failed_count++;
          results.errors.push({
            chunk_id: item.chunk_id,
            error: error.message
          });
        }
      }

      // 全体の成功判定を更新
      results.success = results.failed_count === 0;

      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'EmbeddingManager.importEmbeddingsFromColab',
        error: error,
        severity: 'HIGH',
        context: { file_id: fileId }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }
}
