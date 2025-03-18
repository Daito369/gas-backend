/**
 * RAG 2.0 スプレッドシートストレージ
 * スプレッドシートを使用したデータの保存と取得を管理します
 */
class SheetStorage {

  /**
   * チャンクデータを保存します
   * @param {Object} chunk チャンクデータ
   * @param {string} chunk.id チャンクID
   * @param {string} chunk.document_id ドキュメントID
   * @param {string} chunk.content チャンクコンテンツ
   * @param {Object} chunk.metadata チャンクメタデータ
   * @param {string} chunk.category カテゴリ
   * @return {Object} 保存結果
   */
  static saveChunk(chunk) {
    try {
      // 必須フィールドの検証
      if (!chunk.id || !chunk.document_id || !chunk.content) {
        throw new Error('チャンクの必須フィールドが不足しています');
      }

      // カテゴリを取得（デフォルトは 'general'）
      const category = chunk.category || 'general';

      // メタデータをJSON文字列に変換
      const metadataStr = JSON.stringify(chunk.metadata || {});

      // チャンクの保存先シートを特定
      const sheetInfo = this.getAppropriateChunkSheet(category);

      if (!sheetInfo) {
        throw new Error(`カテゴリ ${category} 用のシートが見つかりません`);
      }

      // チャンクデータの行を作成
      const now = new Date().toISOString();
      const rowData = [
        chunk.id,
        chunk.document_id,
        category,
        chunk.content,
        metadataStr,
        now,
        now
      ];

      // チャンクデータをシートに追加
      const sheet = sheetInfo.sheet;
      sheet.appendRow(rowData);

      return {
        success: true,
        chunk_id: chunk.id,
        sheet_name: sheetInfo.name
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.saveChunk',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_id: chunk.id }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * 埋め込みベクトルを保存します
   * @param {Object} embedding 埋め込みデータ
   * @param {string} embedding.chunk_id チャンクID
   * @param {string} embedding.document_id ドキュメントID
   * @param {Array<number>} embedding.vector 埋め込みベクトル
   * @param {string} embedding.category カテゴリ
   * @param {string} embedding.model_version モデルバージョン
   * @return {Object} 保存結果
   */
  static saveEmbedding(embedding) {
    try {
      // 必須フィールドの検証
      if (!embedding.chunk_id || !embedding.document_id || !embedding.vector) {
        throw new Error('埋め込みの必須フィールドが不足しています');
      }

      // カテゴリを取得（デフォルトは 'general'）
      const category = embedding.category || 'general';

      // 埋め込みベクトルを圧縮してエンコード
      const compressedVector = Utilities.compressAndEncodeArray(embedding.vector);

      // 埋め込みの保存先シートを特定
      const sheetInfo = this.getAppropriateEmbeddingSheet(category);

      if (!sheetInfo) {
        throw new Error(`カテゴリ ${category} 用の埋め込みシートが見つかりません`);
      }

      // 埋め込みデータの行を作成
      const now = new Date().toISOString();
      const modelVersion = embedding.model_version || 'unknown';
      const rowData = [
        embedding.chunk_id,
        embedding.document_id,
        category,
        compressedVector,
        modelVersion,
        now
      ];

      // 埋め込みデータをシートに追加
      const sheet = sheetInfo.sheet;
      sheet.appendRow(rowData);

      return {
        success: true,
        chunk_id: embedding.chunk_id,
        sheet_name: sheetInfo.name
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.saveEmbedding',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_id: embedding.chunk_id }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * ドキュメントのすべてのチャンクを取得します
   * @param {string} documentId ドキュメントID
   * @return {Array<Object>} チャンクの配列
   */
  static getChunksByDocumentId(documentId) {
    try {
      // キャッシュをチェック
      const cacheKey = `doc_chunks_${documentId}`;
      const cachedChunks = CacheManager.get(cacheKey);

      if (cachedChunks) {
        return JSON.parse(cachedChunks);
      }

      // インデックスシートからチャンクの保存場所を取得
      const chunkLocations = this.getDocumentChunkLocations(documentId);
      const chunks = [];

      // 各シートからチャンクを収集
      for (const location of chunkLocations) {
        const sheet = SpreadsheetApp.openById(Config.getDatabaseId()).getSheetByName(location.sheet_name);
        
        if (!sheet) continue;
        
        // シートからデータを取得
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) continue; // ヘッダー行のみの場合はスキップ

        // ヘッダー行を取得してインデックスをマッピング
        const headers = data[0];
        const idIndex = headers.indexOf('id');
        const docIdIndex = headers.indexOf('document_id');
        const categoryIndex = headers.indexOf('category');
        const contentIndex = headers.indexOf('content');
        const metadataIndex = headers.indexOf('metadata');
        const createdAtIndex = headers.indexOf('created_at');
        const updatedAtIndex = headers.indexOf('updated_at');

        // インデックスの検証
        if (idIndex === -1 || docIdIndex === -1 || contentIndex === -1) {
          continue;
        }

        // ドキュメントIDに一致する行を検索
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (row[docIdIndex] === documentId) {
            const chunk = {
              id: row[idIndex],
              document_id: row[docIdIndex],
              content: row[contentIndex]
            };

            // オプションフィールドの追加
            if (categoryIndex !== -1) chunk.category = row[categoryIndex];
            if (metadataIndex !== -1) {
              try {
                chunk.metadata = JSON.parse(row[metadataIndex]);
              } catch (e) {
                chunk.metadata = {};
              }
            }
            if (createdAtIndex !== -1) chunk.created_at = row[createdAtIndex];
            if (updatedAtIndex !== -1) chunk.updated_at = row[updatedAtIndex];

            chunks.push(chunk);
          }
        }
      }

      // 結果をキャッシュに保存 (15分)
      CacheManager.set(cacheKey, JSON.stringify(chunks), 900);

      return chunks;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getChunksByDocumentId',
        error: error,
        severity: 'MEDIUM',
        context: { document_id: documentId }
      });

      return [];
    }
  }

  /**
   * チャンクIDで単一のチャンクを取得します
   * @param {string} chunkId チャンクID
   * @return {Object|null} チャンクオブジェクト、存在しない場合はnull
   */
  static getChunkById(chunkId) {
    try {
      // キャッシュをチェック
      const cacheKey = `chunk_${chunkId}`;
      const cachedChunk = CacheManager.get(cacheKey);

      if (cachedChunk) {
        return JSON.parse(cachedChunk);
      }

      // インデックスシートからチャンクの保存場所を取得
      const chunkLocation = this.getChunkLocation(chunkId);
      
      if (!chunkLocation) {
        return null;
      }
      
      const sheet = SpreadsheetApp.openById(Config.getDatabaseId()).getSheetByName(chunkLocation.sheet_name);
      
      if (!sheet) {
        return null;
      }
      
      // シートからデータを取得
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return null; // ヘッダー行のみの場合はスキップ

      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const idIndex = headers.indexOf('id');

      if (idIndex === -1) return null;

      // チャンクIDに一致する行を検索
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[idIndex] === chunkId) {
          // チャンクオブジェクトを構築
          const chunk = this.rowToChunk(row, headers);

          // キャッシュに結果を保存 (1時間)
          CacheManager.set(cacheKey, JSON.stringify(chunk), 3600);

          return chunk;
        }
      }

      return null;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getChunkById',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_id: chunkId }
      });

      return null;
    }
  }

  /**
   * チャンクIDのリストに基づいてチャンクを取得します
   * @param {Array<string>} chunkIds チャンクIDの配列
   * @return {Array<Object>} チャンクオブジェクトの配列
   */
  static getChunksByIds(chunkIds) {
    try {
      if (!chunkIds || chunkIds.length === 0) {
        return [];
      }

      const results = [];

      // チャンクIDごとに処理
      for (const chunkId of chunkIds) {
        const chunk = this.getChunkById(chunkId);
        if (chunk) {
          results.push(chunk);
        }
      }

      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getChunksByIds',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_ids: chunkIds }
      });

      return [];
    }
  }

  /**
   * チャンクの埋め込みベクトルを取得します
   * @param {string} chunkId チャンクID
   * @return {Array<number>|null} 埋め込みベクトル、存在しない場合はnull
   */
  static getEmbeddingByChunkId(chunkId) {
    try {
      // キャッシュをチェック
      const cacheKey = `emb_${chunkId}`;
      const cachedEmbedding = CacheManager.get(cacheKey);

      if (cachedEmbedding) {
        return JSON.parse(cachedEmbedding);
      }

      // インデックスシートから埋め込みの保存場所を取得
      const embeddingLocation = this.getEmbeddingLocation(chunkId);
      
      if (!embeddingLocation) {
        return null;
      }
      
      const sheet = SpreadsheetApp.openById(Config.getDatabaseId()).getSheetByName(embeddingLocation.sheet_name);
      
      if (!sheet) {
        return null;
      }
      
      // シートからデータを取得
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return null; // ヘッダー行のみの場合はスキップ

      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const chunkIdIndex = headers.indexOf('chunk_id');
      const vectorIndex = headers.indexOf('vector');

      if (chunkIdIndex === -1 || vectorIndex === -1) return null;

      // チャンクIDに一致する行を検索
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[chunkIdIndex] === chunkId) {
          // 圧縮された埋め込みベクトルを復元
          const compressedVector = row[vectorIndex];
          const vector = Utilities.decodeAndDecompressArray(compressedVector);

          // キャッシュに結果を保存 (1時間)
          CacheManager.set(cacheKey, JSON.stringify(vector), 3600);

          return vector;
        }
      }

      return null;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getEmbeddingByChunkId',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_id: chunkId }
      });

      return null;
    }
  }

  /**
   * 複数チャンクの埋め込みベクトルを一括取得します
   * @param {Array<string>} chunkIds チャンクIDの配列
   * @return {Object} チャンクIDをキー、埋め込みベクトルを値とするオブジェクト
   */
  static getEmbeddingsByChunkIds(chunkIds) {
    try {
      if (!chunkIds || chunkIds.length === 0) {
        return {};
      }

      const results = {};

      // チャンクIDをシート別にグループ化して一括取得する最適化
      const locationMap = this.getEmbeddingLocations(chunkIds);
      const sheetGroups = {};
      
      // シート別にチャンクIDをグループ化
      for (const chunkId in locationMap) {
        const location = locationMap[chunkId];
        if (!sheetGroups[location.sheet_name]) {
          sheetGroups[location.sheet_name] = [];
        }
        sheetGroups[location.sheet_name].push(chunkId);
      }
      
      // シートごとに一括取得
      for (const sheetName in sheetGroups) {
        const sheet = SpreadsheetApp.openById(Config.getDatabaseId()).getSheetByName(sheetName);
        if (!sheet) continue;
        
        // シートからデータを取得
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) continue; // ヘッダー行のみの場合はスキップ

        // ヘッダー行を取得してインデックスをマッピング
        const headers = data[0];
        const chunkIdIndex = headers.indexOf('chunk_id');
        const vectorIndex = headers.indexOf('vector');

        if (chunkIdIndex === -1 || vectorIndex === -1) continue;

        // チャンクIDのセットを作成（高速ルックアップのため）
        const chunkIdSet = new Set(sheetGroups[sheetName]);

        // 一致する行を検索
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const currentChunkId = row[chunkIdIndex];

          if (chunkIdSet.has(currentChunkId)) {
            // 圧縮された埋め込みベクトルを復元
            const compressedVector = row[vectorIndex];
            const vector = Utilities.decodeAndDecompressArray(compressedVector);

            // 結果に追加
            results[currentChunkId] = vector;

            // 個別のキャッシュにも保存
            const cacheKey = `emb_${currentChunkId}`;
            CacheManager.set(cacheKey, JSON.stringify(vector), 3600);
          }
        }
      }

      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getEmbeddingsByChunkIds',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_ids_count: chunkIds.length }
      });

      return {};
    }
  }

  /**
   * カテゴリに基づいて検索を実行します
   * @param {string} query 検索クエリ
   * @param {string} category カテゴリ
   * @param {number} limit 最大結果数
   * @return {Array<Object>} 検索結果の配列
   */
  static searchByCategory(query, category, limit = 10) {
    try {
      // カテゴリに関連するすべてのチャンクシートを取得
      const chunkSheets = this.getChunkSheetsByCategory(category);
      const results = [];

      // 単純なキーワード一致検索を実行
      for (const sheetInfo of chunkSheets) {
        const sheet = sheetInfo.sheet;
        const data = sheet.getDataRange().getValues();
        
        // ヘッダー行を取得してインデックスをマッピング
        const headers = data[0];
        const contentIndex = headers.indexOf('content');
        
        if (contentIndex === -1) continue;
        
        // 各行を検索
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const content = row[contentIndex];
          
          // クエリがコンテンツに含まれているか確認
          if (content && content.toLowerCase().includes(query.toLowerCase())) {
            // チャンクオブジェクトを構築して結果に追加
            const chunk = this.rowToChunk(row, headers);
            results.push(chunk);
            
            // 最大結果数に達したら終了
            if (results.length >= limit) {
              break;
            }
          }
        }
        
        // 最大結果数に達したら終了
        if (results.length >= limit) {
          break;
        }
      }

      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.searchByCategory',
        error: error,
        severity: 'MEDIUM',
        context: { query, category }
      });

      return [];
    }
  }

  /**
   * ドキュメントを保存します
   * @param {Object} document ドキュメントデータ
   * @return {Object} 保存結果
   */
  static saveDocument(document) {
    try {
      // 必須フィールドの検証
      if (!document.id || !document.title || !document.content) {
        throw new Error('ドキュメントの必須フィールドが不足しています');
      }

      // データベースを開く
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const documentsSheet = ss.getSheetByName('Documents');

      if (!documentsSheet) {
        throw new Error('Documents シートが見つかりません');
      }

      // メタデータをJSON文字列に変換
      const metadataStr = JSON.stringify(document.metadata || {});

      // カテゴリを取得（デフォルトは 'general'）
      const category = document.category || 'general';

      // ドキュメントデータの行を作成
      const now = new Date().toISOString();
      const rowData = [
        document.id,
        document.title,
        document.path || '',
        category,
        document.language || Utilities.detectLanguage(document.content),
        document.format || 'text',
        metadataStr,
        document.last_updated || now,
        now
      ];

      // 既存のドキュメントを検索して更新または新規追加
      const data = documentsSheet.getDataRange().getValues();
      const idIndex = data[0].indexOf('id');
      
      if (idIndex === -1) {
        throw new Error('Documents シートの形式が不正です');
      }
      
      let rowIndex = -1;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][idIndex] === document.id) {
          rowIndex = i + 1; // シートの行番号は1から始まる
          break;
        }
      }
      
      if (rowIndex > 0) {
        // 既存ドキュメントを更新
        documentsSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        // 新規ドキュメントを追加
        documentsSheet.appendRow(rowData);
      }

      // チャンキングとチャンクの保存
      const chunks = this.chunkDocument(document);
      const chunkResults = chunks.map(chunk => this.saveChunk(chunk));

      return {
        success: true,
        document_id: document.id,
        chunks_count: chunks.length,
        chunks_success: chunkResults.filter(result => result.success).length
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.saveDocument',
        error: error,
        severity: 'MEDIUM',
        context: { document_id: document.id }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * ドキュメントをチャンクに分割します
   * @param {Object} document ドキュメントデータ
   * @return {Array<Object>} チャンクの配列
   */
  static chunkDocument(document) {
    const config = Config.getSystemConfig();
    const chunkSize = config.chunk_size;
    const chunkOverlap = config.chunk_overlap;
    
    // コンテンツをチャンクに分割
    const textChunks = Utilities.splitTextIntoChunks(document.content, chunkSize, chunkOverlap);
    
    // チャンクオブジェクトを作成
    return textChunks.map((chunkText, index) => {
      // メタデータの基本情報をチャンクに引き継ぎ
      const chunkMetadata = {
        ...(document.metadata || {}),
        document_title: document.title,
        chunk_index: index,
        total_chunks: textChunks.length,
        language: document.language || Utilities.detectLanguage(chunkText)
      };
      
      // チャンク固有のメタデータを追加
      const chunkSpecificMetadata = Utilities.extractMetadataFromText(chunkText);
      Object.assign(chunkMetadata, chunkSpecificMetadata);
      
      return {
        id: `${document.id}_chunk_${index}`,
        document_id: document.id,
        content: chunkText,
        category: document.category || 'general',
        metadata: chunkMetadata
      };
    });
  }

  /**
   * ドキュメントデータを取得します
   * @param {string} documentId ドキュメントID
   * @return {Object|null} ドキュメントオブジェクト、存在しない場合はnull
   */
  static getDocumentById(documentId) {
    try {
      // キャッシュをチェック
      const cacheKey = `doc_${documentId}`;
      const cachedDoc = CacheManager.get(cacheKey);

      if (cachedDoc) {
        return JSON.parse(cachedDoc);
      }

      // データベースを開く
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const documentsSheet = ss.getSheetByName('Documents');

      if (!documentsSheet) {
        throw new Error('Documents シートが見つかりません');
      }

      // シートからデータを取得
      const data = documentsSheet.getDataRange().getValues();
      if (data.length <= 1) return null; // ヘッダー行のみの場合はスキップ

      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const idIndex = headers.indexOf('id');
      const titleIndex = headers.indexOf('title');
      const pathIndex = headers.indexOf('path');
      const categoryIndex = headers.indexOf('category');
      const languageIndex = headers.indexOf('language');
      const formatIndex = headers.indexOf('format');
      const metadataIndex = headers.indexOf('metadata');
      const lastUpdatedIndex = headers.indexOf('last_updated');
      const createdAtIndex = headers.indexOf('created_at');

      if (idIndex === -1) return null;

      // ドキュメントIDに一致する行を検索
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[idIndex] === documentId) {
          // ドキュメントオブジェクトを構築
          const document = {
            id: row[idIndex]
          };

          // オプションフィールドの追加
          if (titleIndex !== -1) document.title = row[titleIndex];
          if (pathIndex !== -1) document.path = row[pathIndex];
          if (categoryIndex !== -1) document.category = row[categoryIndex];
          if (languageIndex !== -1) document.language = row[languageIndex];
          if (formatIndex !== -1) document.format = row[formatIndex];
          if (metadataIndex !== -1) {
            try {
              document.metadata = JSON.parse(row[metadataIndex]);
            } catch (e) {
              document.metadata = {};
            }
          }
          if (lastUpdatedIndex !== -1) document.last_updated = row[lastUpdatedIndex];
          if (createdAtIndex !== -1) document.created_at = row[createdAtIndex];

          // キャッシュに結果を保存 (30分)
          CacheManager.set(cacheKey, JSON.stringify(document), 1800);

          return document;
        }
      }

      return null;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getDocumentById',
        error: error,
        severity: 'MEDIUM',
        context: { document_id: documentId }
      });

      return null;
    }
  }

  /**
   * ドキュメントのメタデータを更新します
   * @param {string} documentId ドキュメントID
   * @param {Object} metadata 更新するメタデータ
   * @return {Object} 更新結果
   */
  static updateDocumentMetadata(documentId, metadata) {
    try {
      // データベースを開く
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const documentsSheet = ss.getSheetByName('Documents');

      if (!documentsSheet) {
        throw new Error('Documents シートが見つかりません');
      }

      // シートからデータを取得
      const data = documentsSheet.getDataRange().getValues();
      if (data.length <= 1) {
        throw new Error(`ドキュメント ${documentId} が見つかりません`);
      }

      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const idIndex = headers.indexOf('id');
      const metadataIndex = headers.indexOf('metadata');
      const lastUpdatedIndex = headers.indexOf('last_updated');

      if (idIndex === -1 || metadataIndex === -1) {
        throw new Error('Documents シートの形式が不正です');
      }

      // ドキュメントIDに一致する行を検索
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[idIndex] === documentId) {
          // 現在のメタデータを取得
          let currentMetadata = {};
          try {
            currentMetadata = JSON.parse(row[metadataIndex]);
          } catch (e) {
            // 解析エラーの場合は空オブジェクトを使用
          }

          // メタデータを更新
          const updatedMetadata = { ...currentMetadata, ...metadata };
          const metadataStr = JSON.stringify(updatedMetadata);

          // 行を更新
          const rowIndex = i + 1; // シートの行番号は1から始まる
          documentsSheet.getRange(rowIndex, metadataIndex + 1).setValue(metadataStr);

          // 最終更新日時も更新
          if (lastUpdatedIndex !== -1) {
            const now = new Date().toISOString();
            documentsSheet.getRange(rowIndex, lastUpdatedIndex + 1).setValue(now);
          }

          // キャッシュを削除（次回取得時に更新されたデータを取得するため）
          CacheManager.remove(`doc_${documentId}`);

          return {
            success: true,
            document_id: documentId
          };
        }
      }

      throw new Error(`ドキュメント ${documentId} が見つかりません`);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.updateDocumentMetadata',
        error: error,
        severity: 'MEDIUM',
        context: { document_id: documentId }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * 適切なチャンクシートを取得します
   * @param {string} category カテゴリ
   * @return {Object} シート情報 {name, sheet}
   */
  static getAppropriateChunkSheet(category) {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const indexSheet = ss.getSheetByName('Index_Mapping');
      
      if (!indexSheet) {
        // インデックスシートがない場合は作成
        return this.createNewChunkSheet(category, ss);
      }
      
      const data = indexSheet.getDataRange().getValues();
      
      // ヘッダー行がない場合は作成
      if (data.length === 0) {
        indexSheet.appendRow(['category', 'type', 'sheet_name', 'row_count', 'last_updated']);
        return this.createNewChunkSheet(category, ss);
      }
      
      // カテゴリとタイプに一致するシートを検索
      let candidateSheet = null;
      let candidateSheetName = null;
      let minRowCount = Number.MAX_SAFE_INTEGER;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === category && data[i][1] === 'chunks') {
          const sheetName = data[i][2];
          const rowCount = parseInt(data[i][3], 10) || 0;
          const sheet = ss.getSheetByName(sheetName);
          
          if (sheet && rowCount < minRowCount && rowCount < Config.getSystemConfig().max_rows_per_sheet) {
            candidateSheet = sheet;
            candidateSheetName = sheetName;
            minRowCount = rowCount;
          }
        }
      }
      
      // 適切なシートが見つかった場合はそれを返す
      if (candidateSheet) {
        return { name: candidateSheetName, sheet: candidateSheet };
      }
      
      // 見つからない場合は新しいシートを作成
      return this.createNewChunkSheet(category, ss);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getAppropriateChunkSheet',
        error: error,
        severity: 'HIGH',
        context: { category }
      });
      
      return null;
    }
  }

  /**
   * 新しいチャンクシートを作成します
   * @param {string} category カテゴリ
   * @param {SpreadsheetApp.Spreadsheet} ss スプレッドシート
   * @return {Object} シート情報 {name, sheet}
   */
  static createNewChunkSheet(category, ss) {
    try {
      const indexSheet = ss.getSheetByName('Index_Mapping') || ss.insertSheet('Index_Mapping');
      
      // インデックスシートにヘッダーがない場合は追加
      if (indexSheet.getLastRow() === 0) {
        indexSheet.appendRow(['category', 'type', 'sheet_name', 'row_count', 'last_updated']);
      }
      
      // 新しいシート名を生成
      const shardNumber = this.getNextShardNumber(ss, 'Chunks', category);
      const sheetName = Config.getSheetName('chunks', category, shardNumber);
      
      // シートが既に存在する場合は取得、そうでなければ作成
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        // ヘッダー行を追加
        sheet.appendRow(['id', 'document_id', 'category', 'content', 'metadata', 'created_at', 'updated_at']);
      }
      
      // インデックスシートに追加
      const now = new Date().toISOString();
      indexSheet.appendRow([category, 'chunks', sheetName, 1, now]); // 1はヘッダー行
      
      return { name: sheetName, sheet: sheet };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.createNewChunkSheet',
        error: error,
        severity: 'HIGH',
        context: { category }
      });
      
      return null;
    }
  }

  /**
   * 適切な埋め込みシートを取得します
   * @param {string} category カテゴリ
   * @return {Object} シート情報 {name, sheet}
   */
  static getAppropriateEmbeddingSheet(category) {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const indexSheet = ss.getSheetByName('Index_Mapping');
      
      if (!indexSheet) {
        // インデックスシートがない場合は作成
        return this.createNewEmbeddingSheet(category, ss);
      }
      
      const data = indexSheet.getDataRange().getValues();
      
      // ヘッダー行がない場合は作成
      if (data.length === 0) {
        indexSheet.appendRow(['category', 'type', 'sheet_name', 'row_count', 'last_updated']);
        return this.createNewEmbeddingSheet(category, ss);
      }
      
      // カテゴリとタイプに一致するシートを検索
      let candidateSheet = null;
      let candidateSheetName = null;
      let minRowCount = Number.MAX_SAFE_INTEGER;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === category && data[i][1] === 'embeddings') {
          const sheetName = data[i][2];
          const rowCount = parseInt(data[i][3], 10) || 0;
          const sheet = ss.getSheetByName(sheetName);
          
          if (sheet && rowCount < minRowCount && rowCount < Config.getSystemConfig().max_rows_per_sheet) {
            candidateSheet = sheet;
            candidateSheetName = sheetName;
            minRowCount = rowCount;
          }
        }
      }
      
      // 適切なシートが見つかった場合はそれを返す
      if (candidateSheet) {
        return { name: candidateSheetName, sheet: candidateSheet };
      }
      
      // 見つからない場合は新しいシートを作成
      return this.createNewEmbeddingSheet(category, ss);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getAppropriateEmbeddingSheet',
        error: error,
        severity: 'HIGH',
        context: { category }
      });
      
      return null;
    }
  }

  /**
   * 新しい埋め込みシートを作成します
   * @param {string} category カテゴリ
   * @param {SpreadsheetApp.Spreadsheet} ss スプレッドシート
   * @return {Object} シート情報 {name, sheet}
   */
  static createNewEmbeddingSheet(category, ss) {
    try {
      const indexSheet = ss.getSheetByName('Index_Mapping') || ss.insertSheet('Index_Mapping');
      
      // インデックスシートにヘッダーがない場合は追加
      if (indexSheet.getLastRow() === 0) {
        indexSheet.appendRow(['category', 'type', 'sheet_name', 'row_count', 'last_updated']);
      }
      
      // 新しいシート名を生成
      const shardNumber = this.getNextShardNumber(ss, 'Embeddings', category);
      const sheetName = Config.getSheetName('embeddings', category, shardNumber);
      
      // シートが既に存在する場合は取得、そうでなければ作成
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        // ヘッダー行を追加
        sheet.appendRow(['chunk_id', 'document_id', 'category', 'vector', 'model_version', 'created_at']);
      }
      
      // インデックスシートに追加
      const now = new Date().toISOString();
      indexSheet.appendRow([category, 'embeddings', sheetName, 1, now]); // 1はヘッダー行
      
      return { name: sheetName, sheet: sheet };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.createNewEmbeddingSheet',
        error: error,
        severity: 'HIGH',
        context: { category }
      });
      
      return null;
    }
  }

  /**
   * 次のシャード番号を取得します
   * @param {SpreadsheetApp.Spreadsheet} ss スプレッドシート
   * @param {string} prefix シート名プレフィックス
   * @param {string} category カテゴリ
   * @return {number} 次のシャード番号
   */
  static getNextShardNumber(ss, prefix, category) {
    const sheets = ss.getSheets();
    let maxShardNumber = 0;
    
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      if (sheetName.startsWith(`${prefix}_${category}_`)) {
        // シート名から番号を抽出
        const match = sheetName.match(new RegExp(`${prefix}_${category}_(\\d+)`));
        if (match && match[1]) {
          const shardNumber = parseInt(match[1], 10);
          if (shardNumber > maxShardNumber) {
            maxShardNumber = shardNumber;
          }
        }
      }
    }
    
    return maxShardNumber + 1;
  }

  /**
   * ドキュメントのチャンク保存場所を取得します
   * @param {string} documentId ドキュメントID
   * @return {Array<Object>} チャンク保存場所の配列
   */
  static getDocumentChunkLocations(documentId) {
    try {
      // キャッシュをチェック
      const cacheKey = `doc_chunk_locations_${documentId}`;
      const cachedLocations = CacheManager.get(cacheKey);
      
      if (cachedLocations) {
        return JSON.parse(cachedLocations);
      }
      
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const indexSheet = ss.getSheetByName('Index_Mapping');
      
      if (!indexSheet) {
        return [];
      }
      
      const indexData = indexSheet.getDataRange().getValues();
      if (indexData.length <= 1) {
        return [];
      }
      
      // チャンクタイプのシートを取得
      const chunkSheetNames = [];
      for (let i = 1; i < indexData.length; i++) {
        if (indexData[i][1] === 'chunks') {
          chunkSheetNames.push(indexData[i][2]);
        }
      }
      
      // 各シートでドキュメントIDに関連するチャンクを検索
      const locations = [];
      
      for (const sheetName of chunkSheetNames) {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) continue;
        
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) continue;
        
        const docIdIndex = data[0].indexOf('document_id');
        if (docIdIndex === -1) continue;
        
        // チャンクの有無をチェック（全データを読み込まずに効率化）
        let hasChunks = false;
        for (let i = 1; i < data.length; i++) {
          if (data[i][docIdIndex] === documentId) {
            hasChunks = true;
            break;
          }
        }
        
        if (hasChunks) {
          locations.push({ sheet_name: sheetName });
        }
      }
      
      // キャッシュに保存 (1時間)
      CacheManager.set(cacheKey, JSON.stringify(locations), 3600);
      
      return locations;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getDocumentChunkLocations',
        error: error,
        severity: 'MEDIUM',
        context: { document_id: documentId }
      });
      
      return [];
    }
  }

  /**
   * チャンクの保存場所を取得します
   * @param {string} chunkId チャンクID
   * @return {Object|null} チャンクの保存場所情報
   */
  static getChunkLocation(chunkId) {
    try {
      // キャッシュをチェック
      const cacheKey = `chunk_location_${chunkId}`;
      const cachedLocation = CacheManager.get(cacheKey);
      
      if (cachedLocation) {
        return JSON.parse(cachedLocation);
      }
      
      // ドキュメントIDとチャンクインデックスを抽出
      // チャンクIDの形式: documentId_chunk_index
      const parts = chunkId.split('_chunk_');
      if (parts.length !== 2) {
        return null;
      }
      
      const documentId = parts[0];
      
      // ドキュメントの全チャンク保存場所を取得
      const locations = this.getDocumentChunkLocations(documentId);
      
      if (locations.length === 0) {
        return null;
      }
      
      // 各シートを検索してチャンクを見つける
      for (const location of locations) {
        const ss = SpreadsheetApp.openById(Config.getDatabaseId());
        const sheet = ss.getSheetByName(location.sheet_name);
        
        if (!sheet) continue;
        
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) continue;
        
        const idIndex = data[0].indexOf('id');
        if (idIndex === -1) continue;
        
        // チャンクIDに一致する行を検索
        for (let i = 1; i < data.length; i++) {
          if (data[i][idIndex] === chunkId) {
            const result = { sheet_name: location.sheet_name, row: i + 1 };
            
            // キャッシュに保存 (1時間)
            CacheManager.set(cacheKey, JSON.stringify(result), 3600);
            
            return result;
          }
        }
      }
      
      return null;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getChunkLocation',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_id: chunkId }
      });
      
      return null;
    }
  }

  /**
   * 埋め込みの保存場所を取得します
   * @param {string} chunkId チャンクID
   * @return {Object|null} 埋め込みの保存場所情報
   */
  static getEmbeddingLocation(chunkId) {
    try {
      // キャッシュをチェック
      const cacheKey = `embedding_location_${chunkId}`;
      const cachedLocation = CacheManager.get(cacheKey);
      
      if (cachedLocation) {
        return JSON.parse(cachedLocation);
      }
      
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const indexSheet = ss.getSheetByName('Index_Mapping');
      
      if (!indexSheet) {
        return null;
      }
      
      const indexData = indexSheet.getDataRange().getValues();
      if (indexData.length <= 1) {
        return null;
      }
      
      // 埋め込みタイプのシートを取得
      const embeddingSheetNames = [];
      for (let i = 1; i < indexData.length; i++) {
        if (indexData[i][1] === 'embeddings') {
          embeddingSheetNames.push(indexData[i][2]);
        }
      }
      
      // 各シートでチャンクIDに関連する埋め込みを検索
      for (const sheetName of embeddingSheetNames) {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) continue;
        
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) continue;
        
        const chunkIdIndex = data[0].indexOf('chunk_id');
        if (chunkIdIndex === -1) continue;
        
        // チャンクIDに一致する行を検索
        for (let i = 1; i < data.length; i++) {
          if (data[i][chunkIdIndex] === chunkId) {
            const result = { sheet_name: sheetName, row: i + 1 };
            
            // キャッシュに保存 (1時間)
            CacheManager.set(cacheKey, JSON.stringify(result), 3600);
            
            return result;
          }
        }
      }
      
      return null;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getEmbeddingLocation',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_id: chunkId }
      });
      
      return null;
    }
  }

  /**
   * 複数チャンクの埋め込み保存場所を一括取得します
   * @param {Array<string>} chunkIds チャンクIDの配列
   * @return {Object} チャンクIDをキー、保存場所を値とするオブジェクト
   */
  static getEmbeddingLocations(chunkIds) {
    try {
      if (!chunkIds || chunkIds.length === 0) {
        return {};
      }
      
      const result = {};
      const missingChunkIds = [];
      
      // まずキャッシュから取得を試みる
      for (const chunkId of chunkIds) {
        const cacheKey = `embedding_location_${chunkId}`;
        const cachedLocation = CacheManager.get(cacheKey);
        
        if (cachedLocation) {
          result[chunkId] = JSON.parse(cachedLocation);
        } else {
          missingChunkIds.push(chunkId);
        }
      }
      
      // キャッシュにない場合はスプレッドシートから取得
      if (missingChunkIds.length > 0) {
        const ss = SpreadsheetApp.openById(Config.getDatabaseId());
        const indexSheet = ss.getSheetByName('Index_Mapping');
        
        if (indexSheet) {
          const indexData = indexSheet.getDataRange().getValues();
          
          // 埋め込みタイプのシートを取得
          const embeddingSheetNames = [];
          for (let i = 1; i < indexData.length; i++) {
            if (indexData[i][1] === 'embeddings') {
              embeddingSheetNames.push(indexData[i][2]);
            }
          }
          
          // チャンクIDセットの作成（高速ルックアップのため）
          const chunkIdSet = new Set(missingChunkIds);
          
          // 各シートを検索
          for (const sheetName of embeddingSheetNames) {
            // 全てのチャンクが見つかったら終了
            if (chunkIdSet.size === 0) break;
            
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet) continue;
            
            const data = sheet.getDataRange().getValues();
            if (data.length <= 1) continue;
            
            const chunkIdIndex = data[0].indexOf('chunk_id');
            if (chunkIdIndex === -1) continue;
            
            // 一致する行を検索
            for (let i = 1; i < data.length; i++) {
              const currentChunkId = data[i][chunkIdIndex];
              
              if (chunkIdSet.has(currentChunkId)) {
                const locationInfo = { sheet_name: sheetName, row: i + 1 };
                result[currentChunkId] = locationInfo;
                
                // キャッシュに保存 (1時間)
                const cacheKey = `embedding_location_${currentChunkId}`;
                CacheManager.set(cacheKey, JSON.stringify(locationInfo), 3600);
                
                // 見つかったチャンクIDをセットから削除
                chunkIdSet.delete(currentChunkId);
              }
            }
          }
        }
      }
      
      return result;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getEmbeddingLocations',
        error: error,
        severity: 'MEDIUM',
        context: { chunk_ids_count: chunkIds.length }
      });
      
      return {};
    }
  }

  /**
   * カテゴリに関連するすべてのチャンクシートを取得します
   * @param {string} category カテゴリ（省略時は全カテゴリ）
   * @return {Array<Object>} シート情報の配列 [{name, sheet}]
   */
  static getChunkSheetsByCategory(category) {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const indexSheet = ss.getSheetByName('Index_Mapping');
      
      if (!indexSheet) {
        return [];
      }
      
      const indexData = indexSheet.getDataRange().getValues();
      if (indexData.length <= 1) {
        return [];
      }
      
      const results = [];
      
      // カテゴリに一致するチャンクシートを検索
      for (let i = 1; i < indexData.length; i++) {
        if (indexData[i][1] === 'chunks' && (!category || indexData[i][0] === category)) {
          const sheetName = indexData[i][2];
          const sheet = ss.getSheetByName(sheetName);
          
          if (sheet) {
            results.push({ name: sheetName, sheet: sheet });
          }
        }
      }
      
      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.getChunkSheetsByCategory',
        error: error,
        severity: 'MEDIUM',
        context: { category }
      });
      
      return [];
    }
  }

  /**
   * 行データからチャンクオブジェクトに変換します
   * @param {Array} row 行データ
   * @param {Array<string>} headers ヘッダー
   * @return {Object} チャンクオブジェクト
   */
  static rowToChunk(row, headers) {
    const chunk = {};
    
    // 主要フィールドのインデックスを取得
    const idIndex = headers.indexOf('id');
    const docIdIndex = headers.indexOf('document_id');
    const categoryIndex = headers.indexOf('category');
    const contentIndex = headers.indexOf('content');
    const metadataIndex = headers.indexOf('metadata');
    const createdAtIndex = headers.indexOf('created_at');
    const updatedAtIndex = headers.indexOf('updated_at');
    
    // オブジェクトを構築
    if (idIndex !== -1) chunk.id = row[idIndex];
    if (docIdIndex !== -1) chunk.document_id = row[docIdIndex];
    if (categoryIndex !== -1) chunk.category = row[categoryIndex];
    if (contentIndex !== -1) chunk.content = row[contentIndex];
    if (metadataIndex !== -1) {
      try {
        chunk.metadata = JSON.parse(row[metadataIndex]);
      } catch (e) {
        chunk.metadata = {};
      }
    }
    if (createdAtIndex !== -1) chunk.created_at = row[createdAtIndex];
    if (updatedAtIndex !== -1) chunk.updated_at = row[updatedAtIndex];
    
    return chunk;
  }

  /**
   * 操作を再試行します
   * @param {Object} options 再試行オプション
   * @return {boolean} 再試行が成功したかどうか
   */
  static retryOperation(options) {
    try {
      const { operation, params } = options.context;
      
      // 操作に基づいて適切なメソッドを呼び出す
      switch (operation) {
        case 'saveChunk':
          return this.saveChunk(params).success;
        case 'saveEmbedding':
          return this.saveEmbedding(params).success;
        case 'getChunkById':
          return !!this.getChunkById(params);
        case 'getChunksByDocumentId':
          return !!this.getChunksByDocumentId(params);
        case 'getEmbeddingByChunkId':
          return !!this.getEmbeddingByChunkId(params);
        case 'saveDocument':
          return this.saveDocument(params).success;
        default:
          console.warn(`Unknown operation for retry: ${operation}`);
          return false;
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.retryOperation',
        error: error,
        severity: 'HIGH',
        context: options.context
      });
      
      return false;
    }
  }

  /**
   * バックアップを作成します
   * @param {string} sheetName バックアップするシート名
   * @return {Object} バックアップ結果
   */
  static createBackup(sheetName) {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const sheet = ss.getSheetByName(sheetName);
      
      if (!sheet) {
        throw new Error(`シート ${sheetName} が見つかりません`);
      }
      
      // データを取得
      const data = sheet.getDataRange().getValues();
      
      if (data.length === 0) {
        return { success: true, message: 'バックアップが不要：データなし' };
      }
      
      // ヘッダーと本体データを分離
      const headers = data[0];
      const bodyData = data.slice(1);
      
      // バックアップデータを作成
      const backupData = {
        sheet_name: sheetName,
        headers: headers,
        data: bodyData,
        timestamp: new Date().toISOString()
      };
      
      // バックアップフォルダを取得または作成
      const rootFolder = DriveApp.getFolderById(Config.getRootFolderId());
      let backupFolder;
      const backupFolders = rootFolder.getFoldersByName('Backups');
      
      if (backupFolders.hasNext()) {
        backupFolder = backupFolders.next();
      } else {
        backupFolder = rootFolder.createFolder('Backups');
      }
      
      // サブフォルダを取得または作成（日次/週次）
      const now = new Date();
      const isWeekly = now.getDay() === Config.getSystemConfig().weekly_backup_day;
      
      let targetFolder;
      if (isWeekly) {
        const weeklyFolders = backupFolder.getFoldersByName('Weekly');
        if (weeklyFolders.hasNext()) {
          targetFolder = weeklyFolders.next();
        } else {
          targetFolder = backupFolder.createFolder('Weekly');
        }
      } else {
        const dailyFolders = backupFolder.getFoldersByName('Daily');
        if (dailyFolders.hasNext()) {
          targetFolder = dailyFolders.next();
        } else {
          targetFolder = backupFolder.createFolder('Daily');
        }
      }
      
      // 既存のバックアップを検索して更新または新規作成
      const backupFiles = targetFolder.getFilesByName(`${sheetName}_backup`);
      let backupFile;
      
      if (backupFiles.hasNext()) {
        backupFile = backupFiles.next();
        backupFile.setContent(JSON.stringify(backupData));
      } else {
        backupFile = targetFolder.createFile(`${sheetName}_backup`, JSON.stringify(backupData), 'application/json');
      }
      
      return {
        success: true,
        file_id: backupFile.getId(),
        sheet_name: sheetName,
        is_weekly: isWeekly,
        timestamp: backupData.timestamp
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SheetStorage.createBackup',
        error: error,
        severity: 'HIGH',
        context: { sheet_name: sheetName }
      });
      
      return {
        success: false,
        error: error.message
      };
    }
  }
}
