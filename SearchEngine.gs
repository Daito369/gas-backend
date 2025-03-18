/**
 * RAG 2.0 検索エンジンモジュール
 * ベクトル検索、ハイブリッド検索、検索結果のランキングを管理します
 */
class SearchEngine {

  /**
   * クエリに対する検索を実行します
   * @param {string} query 検索クエリ
   * @param {Object} options 検索オプション
   * @param {string} options.category カテゴリ制限（省略可）
   * @param {string} options.language 言語制限（省略可）
   * @param {number} options.limit 最大結果数（デフォルト: 10）
   * @param {boolean} options.expandQuery クエリ拡張を行うかどうか（デフォルト: true）
   * @param {number} options.semanticWeight セマンティック検索の重み（0.0-1.0、デフォルト: 0.7）
   * @param {number} options.keywordWeight キーワード検索の重み（0.0-1.0、デフォルト: 0.3）
   * @param {boolean} options.useCache キャッシュを使用するかどうか（デフォルト: true）
   * @return {Object} 検索結果 {success: boolean, results: Array, error: string}
   */
  static async search(query, options = {}) {
    try {
      // 開始時間を記録
      const startTime = new Date().getTime();

      // クエリの検証
      if (!query || query.trim().length === 0) {
        throw new Error('検索クエリが空です');
      }

      // デフォルトオプション
      const config = Config.getSystemConfig();
      const defaultOptions = {
        category: '',
        language: '',
        limit: config.max_search_results || 10,
        expandQuery: true,
        semanticWeight: config.semantic_search_weight || 0.7,
        keywordWeight: config.keyword_search_weight || 0.3,
        useCache: true
      };

      // オプションをマージ
      const mergedOptions = { ...defaultOptions, ...options };

      // キャッシュのチェック
      if (mergedOptions.useCache) {
        const cacheKey = this.generateCacheKey(query, mergedOptions);
        const cachedResults = CacheManager.get(cacheKey);
        
        if (cachedResults) {
          const results = JSON.parse(cachedResults);
          
          // キャッシュヒット情報を追加
          results.meta = {
            ...results.meta,
            cache_hit: true,
            timing: {
              ...results.meta?.timing,
              total_ms: new Date().getTime() - startTime
            }
          };
          
          return results;
        }
      }

      // クエリの言語を検出
      const queryLanguage = mergedOptions.language || Utilities.detectLanguage(query);

      // クエリの前処理と拡張
      let processedQuery = query;
      let expandedTerms = [];

      if (mergedOptions.expandQuery) {
        // 言語に応じた前処理
        if (queryLanguage === 'ja') {
          processedQuery = Utilities.preprocessJapaneseText(query);
        } else {
          processedQuery = Utilities.preprocessEnglishText(query);
        }

        // クエリ拡張（同義語、関連語など）
        expandedTerms = await this.expandQuery(processedQuery, queryLanguage);
      }

      // セマンティック検索とキーワード検索を並行実行
      const [semanticResults, keywordResults] = await Promise.all([
        this.performSemanticSearch(processedQuery, {
          category: mergedOptions.category,
          language: queryLanguage,
          limit: Math.min(mergedOptions.limit * 2, 30) // 十分な候補を確保
        }),
        this.performKeywordSearch(processedQuery, expandedTerms, {
          category: mergedOptions.category,
          language: queryLanguage,
          limit: Math.min(mergedOptions.limit * 2, 30) // 十分な候補を確保
        })
      ]);

      // 結果を統合・ランキング
      const combinedResults = this.combineAndRankResults(
        semanticResults, 
        keywordResults, 
        mergedOptions.semanticWeight, 
        mergedOptions.keywordWeight
      );

      // 結果数を制限
      const limitedResults = combinedResults.slice(0, mergedOptions.limit);

      // チャンクからドキュメント情報を取得
      const enhancedResults = await this.enhanceResultsWithDocumentInfo(limitedResults);

      // 結果を整形
      const results = {
        success: true,
        query: query,
        processed_query: processedQuery,
        language: queryLanguage,
        expanded_terms: expandedTerms,
        results: enhancedResults,
        meta: {
          total_count: enhancedResults.length,
          semantic_count: semanticResults.length,
          keyword_count: keywordResults.length,
          timing: {
            total_ms: new Date().getTime() - startTime
          },
          cache_hit: false
        }
      };

      // 結果をキャッシュに保存
      if (mergedOptions.useCache) {
        const cacheKey = this.generateCacheKey(query, mergedOptions);
        const cacheTTL = 600; // 10分
        CacheManager.set(cacheKey, JSON.stringify(results), cacheTTL);
      }

      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.search',
        error: error,
        severity: 'MEDIUM',
        context: { 
          query: query, 
          options: JSON.stringify(options)
        },
        retry: false
      });

      return {
        success: false,
        query: query,
        error: error.message
      };
    }
  }

  /**
   * セマンティック検索を実行します（ベクトル類似度検索）
   * @param {string} query 検索クエリ
   * @param {Object} options 検索オプション
   * @return {Array<Object>} 検索結果の配列
   */
  static async performSemanticSearch(query, options = {}) {
    try {
      // クエリのエンベディングを生成
      const queryEmbeddingResult = await EmbeddingManager.generateQueryEmbedding(query);
      
      if (!queryEmbeddingResult.success || !queryEmbeddingResult.embedding) {
        throw new Error('クエリのエンベディング生成に失敗しました');
      }
      
      const queryEmbedding = queryEmbeddingResult.embedding;
      
      // 検索対象のチャンクシートを決定
      let chunkSheets;
      
      if (options.category) {
        // カテゴリが指定されている場合はそのカテゴリのシートのみ
        chunkSheets = SheetStorage.getChunkSheetsByCategory(options.category);
      } else {
        // 指定がない場合は全カテゴリ
        chunkSheets = SheetStorage.getChunkSheetsByCategory();
      }
      
      if (!chunkSheets || chunkSheets.length === 0) {
        return [];
      }
      
      // 検索結果を保持する配列
      const results = [];
      
      // 各シートに対して検索を実行
      for (const sheetInfo of chunkSheets) {
        const candidates = await this.findSimilarChunksInSheet(
          sheetInfo,
          queryEmbedding,
          options
        );
        
        results.push(...candidates);
      }
      
      // スコアでソート（降順）
      results.sort((a, b) => b.score - a.score);
      
      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.performSemanticSearch',
        error: error,
        severity: 'MEDIUM',
        context: { query: query },
        retry: false
      });
      
      return [];
    }
  }

  /**
   * キーワード検索を実行します
   * @param {string} query 検索クエリ
   * @param {Array<string>} expandedTerms 拡張されたクエリ語
   * @param {Object} options 検索オプション
   * @return {Array<Object>} 検索結果の配列
   */
  static async performKeywordSearch(query, expandedTerms = [], options = {}) {
    try {
      // 検索対象のチャンクシートを決定
      let chunkSheets;
      
      if (options.category) {
        // カテゴリが指定されている場合はそのカテゴリのシートのみ
        chunkSheets = SheetStorage.getChunkSheetsByCategory(options.category);
      } else {
        // 指定がない場合は全カテゴリ
        chunkSheets = SheetStorage.getChunkSheetsByCategory();
      }
      
      if (!chunkSheets || chunkSheets.length === 0) {
        return [];
      }
      
      // 検索キーワードを準備
      let keywords = query.toLowerCase().split(/\s+/).filter(word => word.length > 1);
      
      // 拡張語があれば追加
      if (expandedTerms && expandedTerms.length > 0) {
        keywords = [...keywords, ...expandedTerms];
      }
      
      // 重複を削除
      keywords = [...new Set(keywords)];
      
      // 検索結果を保持する配列
      const results = [];
      
      // 各シートに対して検索を実行
      for (const sheetInfo of chunkSheets) {
        const candidates = await this.findKeywordMatchesInSheet(
          sheetInfo,
          keywords,
          options
        );
        
        results.push(...candidates);
      }
      
      // スコアでソート（降順）
      results.sort((a, b) => b.score - a.score);
      
      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.performKeywordSearch',
        error: error,
        severity: 'MEDIUM',
        context: { 
          query: query,
          expanded_terms_count: expandedTerms ? expandedTerms.length : 0
        },
        retry: false
      });
      
      return [];
    }
  }

  /**
   * シート内で類似チャンクを検索します
   * @param {Object} sheetInfo シート情報 {name, sheet}
   * @param {Array<number>} queryEmbedding クエリのエンベディングベクトル
   * @param {Object} options 検索オプション
   * @return {Array<Object>} 検索結果の配列
   */
  static async findSimilarChunksInSheet(sheetInfo, queryEmbedding, options = {}) {
    try {
      // シートからデータを取得
      const sheet = sheetInfo.sheet;
      const data = sheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        return []; // ヘッダー行のみまたは空のシート
      }
      
      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const idIndex = headers.indexOf('id');
      const documentIdIndex = headers.indexOf('document_id');
      const categoryIndex = headers.indexOf('category');
      const contentIndex = headers.indexOf('content');
      const metadataIndex = headers.indexOf('metadata');
      
      // 必須フィールドの検証
      if (idIndex === -1 || documentIdIndex === -1 || contentIndex === -1) {
        return [];
      }
      
      // チャンク候補と計算結果を保存する配列
      const candidates = [];
      const chunkIds = [];
      
      // 言語フィルターがある場合に使用
      const isLanguageFiltered = !!options.language;
      
      // 各行を処理
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const chunkId = row[idIndex];
        
        // 言語フィルタリング（メタデータから言語をチェック）
        if (isLanguageFiltered && metadataIndex !== -1) {
          let metadata;
          try {
            metadata = JSON.parse(row[metadataIndex] || '{}');
          } catch (e) {
            metadata = {};
          }
          
          // メタデータに言語情報があり、指定された言語と一致しない場合はスキップ
          if (metadata.language && metadata.language !== options.language) {
            continue;
          }
        }
        
        // チャンクIDを記録
        chunkIds.push(chunkId);
        
        // 基本情報を候補に追加
        candidates.push({
          chunk_id: chunkId,
          document_id: row[documentIdIndex],
          category: categoryIndex !== -1 ? row[categoryIndex] : '',
          content: row[contentIndex],
          metadata: metadataIndex !== -1 ? this.parseMetadata(row[metadataIndex]) : {},
          score: 0 // 後で計算
        });
        
        // 候補が多すぎる場合は制限
        if (candidates.length >= (options.limit || 50)) {
          break;
        }
      }
      
      // チャンクIDに対応するエンベディングを取得
      const embeddingsMap = await SheetStorage.getEmbeddingsByChunkIds(chunkIds);
      
      // コサイン類似度を計算して候補のスコアを更新
      for (let i = 0; i < candidates.length; i++) {
        const chunk = candidates[i];
        const embedding = embeddingsMap[chunk.chunk_id];
        
        if (embedding) {
          // コサイン類似度を計算
          chunk.score = this.calculateCosineSimilarity(queryEmbedding, embedding);
        } else {
          // エンベディングがない場合はスコアを0に
          chunk.score = 0;
        }
      }
      
      // スコアでフィルタリング（0より大きいものだけ）
      const filteredCandidates = candidates.filter(chunk => chunk.score > 0);
      
      // スコアでソート（降順）
      filteredCandidates.sort((a, b) => b.score - a.score);
      
      // 上限を適用
      return filteredCandidates.slice(0, options.limit || 20);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.findSimilarChunksInSheet',
        error: error,
        severity: 'MEDIUM',
        context: { sheet_name: sheetInfo.name },
        retry: false
      });
      
      return [];
    }
  }

  /**
   * シート内でキーワード一致を検索します
   * @param {Object} sheetInfo シート情報 {name, sheet}
   * @param {Array<string>} keywords 検索キーワード
   * @param {Object} options 検索オプション
   * @return {Array<Object>} 検索結果の配列
   */
  static async findKeywordMatchesInSheet(sheetInfo, keywords, options = {}) {
    try {
      // シートからデータを取得
      const sheet = sheetInfo.sheet;
      const data = sheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        return []; // ヘッダー行のみまたは空のシート
      }
      
      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const idIndex = headers.indexOf('id');
      const documentIdIndex = headers.indexOf('document_id');
      const categoryIndex = headers.indexOf('category');
      const contentIndex = headers.indexOf('content');
      const metadataIndex = headers.indexOf('metadata');
      
      // 必須フィールドの検証
      if (idIndex === -1 || documentIdIndex === -1 || contentIndex === -1) {
        return [];
      }
      
      // 検索結果を保持する配列
      const results = [];
      
      // 言語フィルターがある場合に使用
      const isLanguageFiltered = !!options.language;
      
      // 各行を処理
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const content = (row[contentIndex] || '').toLowerCase();
        
        // 言語フィルタリング（メタデータから言語をチェック）
        if (isLanguageFiltered && metadataIndex !== -1) {
          let metadata;
          try {
            metadata = JSON.parse(row[metadataIndex] || '{}');
          } catch (e) {
            metadata = {};
          }
          
          // メタデータに言語情報があり、指定された言語と一致しない場合はスキップ
          if (metadata.language && metadata.language !== options.language) {
            continue;
          }
        }
        
        // キーワード一致を検索
        const matchedKeywords = keywords.filter(keyword => content.includes(keyword.toLowerCase()));
        
        // 一致したキーワードがある場合のみ結果に追加
        if (matchedKeywords.length > 0) {
          // キーワード一致のスコアを計算（BM25アルゴリズムの簡易版）
          // 一致数とキーワードの重要度に基づくスコア
          const score = this.calculateKeywordMatchScore(content, matchedKeywords);
          
          results.push({
            chunk_id: row[idIndex],
            document_id: row[documentIdIndex],
            category: categoryIndex !== -1 ? row[categoryIndex] : '',
            content: row[contentIndex],
            metadata: metadataIndex !== -1 ? this.parseMetadata(row[metadataIndex]) : {},
            matched_keywords: matchedKeywords,
            score: score
          });
        }
        
        // 結果が多すぎる場合は制限
        if (results.length >= (options.limit || 50)) {
          break;
        }
      }
      
      // スコアでソート（降順）
      results.sort((a, b) => b.score - a.score);
      
      // 上限を適用
      return results.slice(0, options.limit || 20);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.findKeywordMatchesInSheet',
        error: error,
        severity: 'MEDIUM',
        context: { 
          sheet_name: sheetInfo.name,
          keywords: keywords.join(',')
        },
        retry: false
      });
      
      return [];
    }
  }

  /**
   * セマンティック検索とキーワード検索の結果を統合してランキングします
   * @param {Array<Object>} semanticResults セマンティック検索結果
   * @param {Array<Object>} keywordResults キーワード検索結果
   * @param {number} semanticWeight セマンティック検索の重み（0.0-1.0）
   * @param {number} keywordWeight キーワード検索の重み（0.0-1.0）
   * @return {Array<Object>} 統合された検索結果
   */
  static combineAndRankResults(semanticResults, keywordResults, semanticWeight = 0.7, keywordWeight = 0.3) {
    try {
      // 重みの正規化
      const totalWeight = semanticWeight + keywordWeight;
      const normalizedSemanticWeight = semanticWeight / totalWeight;
      const normalizedKeywordWeight = keywordWeight / totalWeight;
      
      // すべてのチャンクIDを収集
      const allChunkIds = new Set([
        ...semanticResults.map(item => item.chunk_id),
        ...keywordResults.map(item => item.chunk_id)
      ]);
      
      // 結果を統合
      const combinedResults = [];
      
      // チャンクIDをマップに変換して高速ルックアップ
      const semanticMap = new Map(semanticResults.map(item => [item.chunk_id, item]));
      const keywordMap = new Map(keywordResults.map(item => [item.chunk_id, item]));
      
      // 各チャンクに対して統合スコアを計算
      for (const chunkId of allChunkIds) {
        const semanticResult = semanticMap.get(chunkId);
        const keywordResult = keywordMap.get(chunkId);
        
        // どちらかの結果から基本情報を取得
        const baseResult = semanticResult || keywordResult;
        
        // スコアを計算
        let combinedScore = 0;
        
        // セマンティックスコアを考慮
        if (semanticResult) {
          combinedScore += semanticResult.score * normalizedSemanticWeight;
        }
        
        // キーワードスコアを考慮
        if (keywordResult) {
          combinedScore += keywordResult.score * normalizedKeywordWeight;
        }
        
        // 統合結果に追加
        combinedResults.push({
          ...baseResult,
          score: combinedScore,
          semantic_score: semanticResult ? semanticResult.score : 0,
          keyword_score: keywordResult ? keywordResult.score : 0,
          matched_keywords: keywordResult ? keywordResult.matched_keywords : []
        });
      }
      
      // スコアでソート（降順）
      combinedResults.sort((a, b) => b.score - a.score);
      
      return combinedResults;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.combineAndRankResults',
        error: error,
        severity: 'MEDIUM',
        context: { 
          semantic_count: semanticResults.length,
          keyword_count: keywordResults.length
        },
        retry: false
      });
      
      // エラーの場合はセマンティック結果を優先
      return semanticResults;
    }
  }

  /**
   * 結果にドキュメント情報を追加して強化します
   * @param {Array<Object>} results 検索結果の配列
   * @return {Array<Object>} 強化された検索結果
   */
  static async enhanceResultsWithDocumentInfo(results) {
    try {
      if (!results || results.length === 0) {
        return [];
      }
      
      // ドキュメントIDを収集（重複を除去）
      const documentIds = [...new Set(results.map(result => result.document_id))];
      
      // ドキュメント情報をキャッシュに記録
      const documentInfoMap = {};
      
      // 各ドキュメントの情報を取得
      for (const documentId of documentIds) {
        const document = SheetStorage.getDocumentById(documentId);
        
        if (document) {
          documentInfoMap[documentId] = {
            title: document.title || 'Untitled',
            path: document.path || '',
            format: document.format || 'unknown',
            language: document.language || 'unknown',
            category: document.category || 'general',
            last_updated: document.last_updated || '',
            metadata: document.metadata || {}
          };
        }
      }
      
      // 結果を強化
      const enhancedResults = results.map(result => {
        const docInfo = documentInfoMap[result.document_id] || {};
        
        // スニペットを抽出（内容の最初の200文字程度）
        const snippet = this.extractSnippet(result.content, result.matched_keywords);
        
        return {
          ...result,
          title: docInfo.title || 'Untitled',
          path: docInfo.path || '',
          format: docInfo.format || 'unknown',
          language: docInfo.language || result.metadata?.language || 'unknown',
          category: docInfo.category || result.category || 'general',
          last_updated: docInfo.last_updated || '',
          document_metadata: docInfo.metadata || {},
          snippet: snippet,
          relevance_score: result.score
        };
      });
      
      return enhancedResults;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.enhanceResultsWithDocumentInfo',
        error: error,
        severity: 'MEDIUM',
        context: { results_count: results.length },
        retry: false
      });
      
      return results;
    }
  }

  /**
   * クエリを拡張します（同義語、関連語などを追加）
   * @param {string} query 検索クエリ
   * @param {string} language 言語
   * @return {Array<string>} 拡張語の配列
   */
  static async expandQuery(query, language) {
    try {
      // 言語に応じて拡張方法を変更
      if (language === 'ja') {
        return this.expandJapaneseQuery(query);
      } else {
        return this.expandEnglishQuery(query);
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.expandQuery',
        error: error,
        severity: 'LOW',
        context: { query, language },
        retry: false
      });
      
      return [];
    }
  }

  /**
   * 日本語クエリを拡張します
   * @param {string} query 検索クエリ
   * @return {Array<string>} 拡張語の配列
   */
  static async expandJapaneseQuery(query) {
    // 簡易的な日本語キーワード拡張（本番環境では辞書やGemini APIを活用）
    const expansions = [];
    
    // 基本的な同義語マッピング
    const synonyms = {
      '広告': ['アド', 'プロモーション', '宣伝'],
      '検索': ['サーチ', '探す', '探索'],
      'キャンペーン': ['施策', 'プロモーション'],
      '予算': ['コスト', '費用', '金額'],
      '入札': ['ビッド', '入札価格', '掲載順位'],
      '表示': ['インプレッション', '露出', '表出'],
      'クリック': ['タップ', 'タッチ'],
      'コンバージョン': ['CV', '成約', '目標達成'],
      '最適化': ['パフォーマンス向上', '改善'],
      '設定': ['セットアップ', '構成', 'コンフィグ'],
      'レポート': ['分析', '統計', 'データ'],
      'アカウント': ['アド アカウント', 'ユーザー'],
      'アフィリエイト': ['パートナー', '提携'],
      'モバイル': ['スマホ', 'スマートフォン', '携帯'],
      'ランディングページ': ['LP', '着地ページ'],
      'キーワード': ['検索語句', '検索クエリ'],
      '掲載順位': ['ランク', '位置', '表示位置'],
      'インプレッション': ['表示回数', '露出', '広告表示'],
      '最適化スコア': ['オプティマイズ スコア'],
      'レポート': ['リポート', 'データ', '分析'],
      'モバイル': ['スマホ', 'スマートフォン', '携帯']
    };
    
    // キーワードを分解
    const keywords = query.split(/[\s\u3000]+/);
    
    // 各キーワードに対して同義語を追加
    for (const keyword of keywords) {
      // 完全一致の同義語
      if (synonyms[keyword]) {
        expansions.push(...synonyms[keyword]);
      }
      
      // 部分一致の同義語
      for (const [term, synonymList] of Object.entries(synonyms)) {
        if (keyword.includes(term)) {
          expansions.push(...synonymList);
        } else if (term.includes(keyword) && keyword.length > 1) {
          expansions.push(term);
        }
      }
    }
    
    // 重複を削除して返す
    return [...new Set(expansions)];
  }

  /**
   * 英語クエリを拡張します
   * @param {string} query 検索クエリ
   * @return {Array<string>} 拡張語の配列
   */
  static async expandEnglishQuery(query) {
    // 簡易的な英語キーワード拡張（本番環境では辞書やGemini APIを活用）
    const expansions = [];
    
    // 基本的な同義語マッピング
    const synonyms = {
      'ad': ['advertisement', 'promotion', 'advert'],
      'search': ['find', 'query', 'lookup'],
      'campaign': ['promotion', 'initiative', 'effort'],
      'budget': ['cost', 'spend', 'funding', 'finance'],
      'bid': ['offer', 'auction', 'price'],
      'impression': ['view', 'display', 'exposure'],
      'click': ['tap', 'press', 'selection'],
      'conversion': ['cv', 'goal completion', 'acquisition'],
      'optimization': ['improvement', 'enhancement', 'refinement'],
      'setup': ['configuration', 'settings', 'establish'],
      'report': ['analysis', 'statistics', 'data', 'metrics'],
      'account': ['profile', 'ads account', 'user'],
      'affiliate': ['partner', 'associate', 'referral'],
      'mobile': ['smartphone', 'handheld', 'cell phone'],
      'landing page': ['lp', 'destination page', 'target page'],
      'keyword': ['search term', 'query term', 'search query'],
      'position': ['rank', 'placement', 'listing'],
      'impression': ['view', 'display', 'appearance'],
      'optimization score': ['optimize score', 'performance rating'],
      'report': ['reporting', 'analytics', 'data', 'statistics'],
      'audience': ['target group', 'demographic', 'viewers']
    };
    
    // キーワードを分解
    const keywords = query.toLowerCase().split(/\s+/);
    
    // 各キーワードに対して同義語を追加
    for (const keyword of keywords) {
      // 完全一致の同義語
      if (synonyms[keyword]) {
        expansions.push(...synonyms[keyword]);
      }
      
      // 部分一致の同義語
      for (const [term, synonymList] of Object.entries(synonyms)) {
        if (keyword.includes(term)) {
          expansions.push(...synonymList);
        } else if (term.includes(keyword) && keyword.length > 2) {
          expansions.push(term);
        }
      }
    }
    
    // 重複を削除して返す
    return [...new Set(expansions)];
  }

  /**
   * コサイン類似度を計算します
   * @param {Array<number>} vecA ベクトルA
   * @param {Array<number>} vecB ベクトルB
   * @return {number} コサイン類似度（0～1）
   */
  static calculateCosineSimilarity(vecA, vecB) {
    try {
      // ベクトルの次元数が一致しない場合
      if (vecA.length !== vecB.length) {
        return 0;
      }
      
      // ドット積を計算
      let dotProduct = 0;
      let normA = 0;
      let normB = 0;
      
      for (let i = 0; i < vecA.length; i++) {
        dotProduct += vecA[i] * vecB[i];
        normA += vecA[i] * vecA[i];
        normB += vecB[i] * vecB[i];
      }
      
      // ノルムを計算
      normA = Math.sqrt(normA);
      normB = Math.sqrt(normB);
      
      // ゼロ除算防止
      if (normA === 0 || normB === 0) {
        return 0;
      }
      
      // コサイン類似度を計算して返す
      return dotProduct / (normA * normB);
    } catch (error) {
      return 0;
    }
  }

  /**
   * キーワード一致スコアを計算します
   * @param {string} content 検索対象テキスト
   * @param {Array<string>} matchedKeywords 一致したキーワード
   * @return {number} キーワード一致スコア
   */
  static calculateKeywordMatchScore(content, matchedKeywords) {
    try {
      if (!matchedKeywords || matchedKeywords.length === 0) {
        return 0;
      }
      
      const contentLower = content.toLowerCase();
      let score = 0;
      
      // 各キーワードに対してスコアを計算
      for (const keyword of matchedKeywords) {
        // 出現回数をカウント
        let count = 0;
        let pos = contentLower.indexOf(keyword.toLowerCase());
        
        while (pos !== -1) {
          count++;
          pos = contentLower.indexOf(keyword.toLowerCase(), pos + 1);
        }
        
        // キーワードの長さに応じた重み付け（長いキーワードほど重要）
        const lengthWeight = Math.sqrt(keyword.length) / 2;
        
        // 位置に基づく重み付け（先頭に近いほど重要）
        const firstPosition = contentLower.indexOf(keyword.toLowerCase());
        const positionWeight = firstPosition < 100 ? 1.5 : 
                              firstPosition < 300 ? 1.2 : 1.0;
        
        // 出現回数に基づくスコア（TF-IDFの簡易版）
        // 回数が多いほど重要だが、diminishing returns
        const frequencyScore = Math.sqrt(count) * 0.5;
        
        // キーワードごとのスコアを加算
        score += (1 + frequencyScore) * lengthWeight * positionWeight;
      }
      
      // キーワードの多様性に基づくボーナス
      const diversityBonus = Math.sqrt(matchedKeywords.length) * 0.2;
      
      // 最終スコアを計算
      return score * (1 + diversityBonus);
    } catch (error) {
      return matchedKeywords.length * 0.5; // エラー時の保守的な推定
    }
  }

  /**
   * JSONメタデータを解析します
   * @param {string} metadataStr JSONメタデータ文字列
   * @return {Object} メタデータオブジェクト
   */
  static parseMetadata(metadataStr) {
    try {
      return JSON.parse(metadataStr || '{}');
    } catch (e) {
      return {};
    }
  }

  /**
   * 検索クエリに基づいてキャッシュキーを生成します
   * @param {string} query 検索クエリ
   * @param {Object} options 検索オプション
   * @return {string} キャッシュキー
   */
  static generateCacheKey(query, options) {
    const normalizedQuery = query.trim().toLowerCase();
    const optionsHash = JSON.stringify({
      category: options.category || '',
      language: options.language || '',
      limit: options.limit || 10,
      semanticWeight: options.semanticWeight || 0.7,
      keywordWeight: options.keywordWeight || 0.3
    });
    
    return `search_${normalizedQuery}_${Utilities.generateUniqueId().substring(0, 8)}_${optionsHash}`;
  }

  /**
   * コンテンツからスニペットを抽出します
   * @param {string} content コンテンツ
   * @param {Array<string>} keywords キーワード（ハイライト用）
   * @param {number} maxLength 最大長さ（デフォルト: 200）
   * @return {string} スニペット
   */
  static extractSnippet(content, keywords = [], maxLength = 200) {
    try {
      if (!content) {
        return '';
      }
      
      // コンテンツの長さが最大長さ以下の場合はそのまま返す
      if (content.length <= maxLength) {
        return content;
      }
      
      // キーワードがある場合は最初のキーワードの周辺を抽出
      if (keywords && keywords.length > 0) {
        for (const keyword of keywords) {
          const keywordLower = keyword.toLowerCase();
          const contentLower = content.toLowerCase();
          const index = contentLower.indexOf(keywordLower);
          
          if (index !== -1) {
            // キーワードの前後の文脈を含むスニペット
            const startPos = Math.max(0, index - 80);
            const endPos = Math.min(content.length, index + keyword.length + 120);
            
            let snippet = content.substring(startPos, endPos);
            
            // スニペットの前後に「...」を追加（必要な場合）
            if (startPos > 0) {
              snippet = '...' + snippet;
            }
            
            if (endPos < content.length) {
              snippet = snippet + '...';
            }
            
            return snippet;
          }
        }
      }
      
      // キーワードがない場合や一致しない場合は先頭から抽出
      return content.substring(0, maxLength) + '...';
    } catch (error) {
      return content ? content.substring(0, maxLength) + '...' : '';
    }
  }

  /**
   * カテゴリ一覧を取得します
   * @return {Array<string>} カテゴリの配列
   */
  static getCategories() {
    try {
      // スプレッドシートからカテゴリ一覧を取得
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const indexSheet = ss.getSheetByName('Index_Mapping');
      
      if (!indexSheet) {
        return this.getDefaultCategories();
      }
      
      const data = indexSheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        return this.getDefaultCategories();
      }
      
      // カテゴリ列（通常は1列目）から一意のカテゴリを抽出
      const categories = new Set();
      
      for (let i = 1; i < data.length; i++) {
        const category = data[i][0];
        if (category) {
          categories.add(category);
        }
      }
      
      return [...categories];
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.getCategories',
        error: error,
        severity: 'LOW',
        retry: false
      });
      
      return this.getDefaultCategories();
    }
  }

  /**
   * デフォルトのカテゴリ一覧を返します
   * @return {Array<string>} デフォルトカテゴリの配列
   */
  static getDefaultCategories() {
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
        case 'search':
          const searchResult = await this.search(params.query, params.options);
          return searchResult.success;
        
        case 'performSemanticSearch':
          const semanticResults = await this.performSemanticSearch(params.query, params.options);
          return semanticResults.length > 0;
        
        case 'performKeywordSearch':
          const keywordResults = await this.performKeywordSearch(params.query, params.expandedTerms, params.options);
          return keywordResults.length > 0;
        
        default:
          console.warn(`Unknown operation for retry: ${operation}`);
          return false;
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'SearchEngine.retryOperation',
        error: error,
        severity: 'HIGH',
        context: options.context
      });
      
      return false;
    }
  }
}
