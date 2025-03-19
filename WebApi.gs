/**
 * RAG 2.0 Web API モジュール
 * フロントエンドとバックエンド間の通信を管理するRESTful APIを提供します
 */
class WebApi {

  /**
   * API処理の入口関数
   * @param {Object} e イベントオブジェクト
   * @return {Object} APIレスポンス
   */
  static doPost(e) {
    try {
      // リクエストタイムアウト設定
      const timeout = setTimeout(() => {
        throw new Error('API request timeout');
      }, Config.getSystemConfig().api_timeout || 30000);
      
      // パラメータとプロパティの取得
      const params = e.parameter;
      let payload = {};
      
      // POSTデータがある場合はJSONとして解析
      if (e.postData && e.postData.contents) {
        try {
          payload = JSON.parse(e.postData.contents);
        } catch (error) {
          return this.createErrorResponse('Invalid JSON in request body', 400);
        }
      }
      
      // クエリパラメータとJSONペイロードをマージ
      const requestData = { ...params, ...payload };
      
      // リクエストタイプの検証
      if (!requestData.type) {
        return this.createErrorResponse('Request type is required', 400);
      }
      
      // APIキーを取得（オプション）
      const apiKey = requestData.api_key || '';
      
      // リクエストタイプに基づいてハンドラーを呼び出す
      let result;
      switch (requestData.type) {
        case 'search':
          result = this.handleSearchRequest(requestData, apiKey);
          break;
          
        case 'generate_response':
          result = this.handleGenerateResponseRequest(requestData, apiKey);
          break;
          
        case 'get_document':
          result = this.handleGetDocumentRequest(requestData, apiKey);
          break;
          
        case 'get_templates':
          result = this.handleGetTemplatesRequest(requestData, apiKey);
          break;
          
        case 'get_categories':
          result = this.handleGetCategoriesRequest(requestData, apiKey);
          break;
          
        case 'process_document':
          result = this.handleProcessDocumentRequest(requestData, apiKey);
          break;
          
        case 'health_check':
          result = this.handleHealthCheckRequest(requestData, apiKey);
          break;
          
        default:
          result = this.createErrorResponse(`Unknown request type: ${requestData.type}`, 400);
      }
      
      // タイムアウトをクリア
      clearTimeout(timeout);
      
      return result;
    } catch (error) {
      // エラーハンドリング
      ErrorHandler.handleError({
        source: 'WebApi.doPost',
        error: error,
        severity: 'HIGH',
        context: { request_type: e?.parameter?.type || 'unknown' }
      });
      
      return this.createErrorResponse(error.message, 500);
    }
  }

  /**
   * APIフロントエンドページを提供します
   * @param {Object} e イベントオブジェクト
   * @return {HtmlOutput} HTMLレスポンス
   */
  static doGet(e) {
    try {
      // パラメータの取得
      const params = e.parameter;
      const view = params.view || 'index';
      
      // リクエストされたビューに基づいてHTMLを返す
      let htmlContent;
      switch (view) {
        case 'index':
          htmlContent = HtmlService.createHtmlOutputFromFile('index')
            .setTitle('Google広告サポート RAG 2.0');
          break;
          
        case 'styles':
          return ContentService.createTextOutput()
            .setMimeType(ContentService.MimeType.CSS)
            .setContent(this.getStylesContent());
          
        case 'script':
          return ContentService.createTextOutput()
            .setMimeType(ContentService.MimeType.JAVASCRIPT)
            .setContent(this.getScriptContent());
          
        case 'components':
          htmlContent = HtmlService.createHtmlOutputFromFile('components')
            .setTitle('Components');
          break;
          
        case 'config':
          // フロントエンド設定を返す（認証が必要）
          if (!this.isUserAuthenticated(e)) {
            return HtmlService.createHtmlOutput('Authentication required')
              .setTitle('Error');
          }
          
          return ContentService.createTextOutput()
            .setMimeType(ContentService.MimeType.JSON)
            .setContent(JSON.stringify(this.getFrontendConfig()));
          
        default:
          htmlContent = HtmlService.createHtmlOutput('<h1>Page Not Found</h1>')
            .setTitle('Error');
      }
      
      // XSSプロテクションヘッダーを追加
      htmlContent.addHeader('X-XSS-Protection', '1; mode=block');
      htmlContent.addHeader('X-Content-Type-Options', 'nosniff');
      
      return htmlContent;
    } catch (error) {
      // エラーハンドリング
      ErrorHandler.handleError({
        source: 'WebApi.doGet',
        error: error,
        severity: 'MEDIUM',
        context: { view: e?.parameter?.view || 'index' }
      });
      
      return HtmlService.createHtmlOutput(`<h1>Error</h1><p>${error.message}</p>`)
        .setTitle('Error');
    }
  }

  /**
   * 検索リクエストを処理します
   * @param {Object} requestData リクエストデータ
   * @param {string} apiKey APIキー（オプション）
   * @return {Object} APIレスポンス
   */
  static async handleSearchRequest(requestData, apiKey) {
    try {
      // 必須フィールドの検証
      if (!requestData.query) {
        return this.createErrorResponse('Query is required', 400);
      }
      
      // 検索オプションの構築
      const options = {
        category: requestData.category || '',
        language: requestData.language || '',
        limit: parseInt(requestData.limit) || 10,
        expandQuery: requestData.expand_query !== 'false',
        semanticWeight: parseFloat(requestData.semantic_weight) || 0.7,
        keywordWeight: parseFloat(requestData.keyword_weight) || 0.3,
        useCache: requestData.use_cache !== 'false'
      };
      
      // APIキーがある場合はGemini APIを設定
      if (apiKey) {
        GeminiIntegration.setUserApiKey(apiKey);
      }
      
      // 検索を実行
      const searchResults = await SearchEngine.search(requestData.query, options);
      
      // リクエスト制限を回避するためにレスポンスサイズを最適化
      const optimizedResults = this.optimizeSearchResults(searchResults);
      
      return this.createSuccessResponse(optimizedResults);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.handleSearchRequest',
        error: error,
        severity: 'MEDIUM',
        context: { query: requestData.query }
      });
      
      return this.createErrorResponse(error.message, 500);
    }
  }

  /**
   * 応答生成リクエストを処理します
   * @param {Object} requestData リクエストデータ
   * @param {string} apiKey APIキー（オプション）
   * @return {Object} APIレスポンス
   */
  static async handleGenerateResponseRequest(requestData, apiKey) {
    try {
      // 必須フィールドの検証
      if (!requestData.query) {
        return this.createErrorResponse('Query is required', 400);
      }
      
      if (!requestData.search_results && !requestData.search_results_id) {
        return this.createErrorResponse('Search results or search_results_id is required', 400);
      }
      
      // 検索結果の取得
      let searchResults;
      if (requestData.search_results) {
        // 直接検索結果が提供された場合
        searchResults = requestData.search_results;
      } else if (requestData.search_results_id) {
        // キャッシュから検索結果を取得
        const cachedResults = CacheManager.get(`search_results_${requestData.search_results_id}`);
        if (!cachedResults) {
          return this.createErrorResponse('Search results not found in cache', 404);
        }
        searchResults = JSON.parse(cachedResults);
      }
      
      // 応答生成オプションの構築
      const options = {
        responseType: requestData.response_type || 'standard',
        language: requestData.language || searchResults.language || 'ja',
        templateId: requestData.template_id || null,
        customParams: requestData.custom_params || {},
        enhanceWithGemini: requestData.enhance_with_gemini !== 'false'
      };
      
      // APIキーがある場合はGemini APIを設定
      if (apiKey) {
        GeminiIntegration.setUserApiKey(apiKey);
      }
      
      // 応答を生成
      const response = await ResponseGenerator.generateResponse(searchResults, requestData.query, options);
      
      return this.createSuccessResponse(response);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.handleGenerateResponseRequest',
        error: error,
        severity: 'MEDIUM',
        context: { query: requestData.query, response_type: requestData.response_type }
      });
      
      return this.createErrorResponse(error.message, 500);
    }
  }

  /**
   * ドキュメント取得リクエストを処理します
   * @param {Object} requestData リクエストデータ
   * @param {string} apiKey APIキー（オプション）
   * @return {Object} APIレスポンス
   */
  static handleGetDocumentRequest(requestData, apiKey) {
    try {
      // 必須フィールドの検証
      if (!requestData.document_id) {
        return this.createErrorResponse('Document ID is required', 400);
      }
      
      // ドキュメントを取得
      const document = SheetStorage.getDocumentById(requestData.document_id);
      
      if (!document) {
        return this.createErrorResponse('Document not found', 404);
      }
      
      // 言語パラメータに基づいて翻訳されたバージョンを取得
      if (requestData.language && document.language !== requestData.language) {
        const translatedDoc = MultilangProcessor.getDocumentInLanguage(
          requestData.document_id, 
          requestData.language
        );
        
        if (translatedDoc) {
          return this.createSuccessResponse(translatedDoc);
        }
      }
      
      return this.createSuccessResponse(document);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.handleGetDocumentRequest',
        error: error,
        severity: 'MEDIUM',
        context: { document_id: requestData.document_id }
      });
      
      return this.createErrorResponse(error.message, 500);
    }
  }

  /**
   * テンプレート取得リクエストを処理します
   * @param {Object} requestData リクエストデータ
   * @param {string} apiKey APIキー（オプション）
   * @return {Object} APIレスポンス
   */
  static async handleGetTemplatesRequest(requestData, apiKey) {
    try {
      // テンプレートタイプによるフィルタリング
      const templateType = requestData.template_type || '';
      const language = requestData.language || '';
      
      // テンプレートをロード
      const templates = await ResponseGenerator.loadTemplates();
      
      // フィルタリング
      let filteredTemplates = templates;
      
      if (templateType) {
        filteredTemplates = filteredTemplates.filter(template => template.type === templateType);
      }
      
      if (language) {
        filteredTemplates = filteredTemplates.filter(template => !template.language || template.language === language);
      }
      
      // レスポンスサイズを最適化（コンテンツを除外）
      const optimizedTemplates = filteredTemplates.map(template => ({
        id: template.id,
        name: template.name,
        type: template.type,
        language: template.language,
        category: template.category,
        metadata: template.metadata
      }));
      
      return this.createSuccessResponse(optimizedTemplates);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.handleGetTemplatesRequest',
        error: error,
        severity: 'MEDIUM',
        context: { template_type: requestData.template_type }
      });
      
      return this.createErrorResponse(error.message, 500);
    }
  }

  /**
   * カテゴリ取得リクエストを処理します
   * @param {Object} requestData リクエストデータ
   * @param {string} apiKey APIキー（オプション）
   * @return {Object} APIレスポンス
   */
  static handleGetCategoriesRequest(requestData, apiKey) {
    try {
      // カテゴリを取得
      const categories = SearchEngine.getCategories();
      
      return this.createSuccessResponse(categories);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.handleGetCategoriesRequest',
        error: error,
        severity: 'LOW'
      });
      
      return this.createErrorResponse(error.message, 500);
    }
  }

  /**
   * ドキュメント処理リクエストを処理します
   * @param {Object} requestData リクエストデータ
   * @param {string} apiKey APIキー（オプション）
   * @return {Object} APIレスポンス
   */
  static handleProcessDocumentRequest(requestData, apiKey) {
    try {
      // 認証チェック（管理者機能のため）
      if (!this.isAdminUser(apiKey)) {
        return this.createErrorResponse('Admin access required', 403);
      }
      
      // 必須フィールドの検証
      if (!requestData.file_id) {
        return this.createErrorResponse('File ID is required', 400);
      }
      
      // 処理オプションの構築
      const options = {
        language: requestData.language || '',
        category: requestData.category || 'general',
        generateEmbeddings: requestData.generate_embeddings === 'true'
      };
      
      // ドキュメントを処理
      const result = DocumentProcessor.processDocument(requestData.file_id, options);
      
      if (!result.success) {
        return this.createErrorResponse(result.error || 'Document processing failed', 500);
      }
      
      // 埋め込みを生成（オプション）
      if (options.generateEmbeddings && result.document_id) {
        // 非同期で埋め込みを生成
        this.scheduleEmbeddingGeneration(result.document_id);
      }
      
      return this.createSuccessResponse(result);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.handleProcessDocumentRequest',
        error: error,
        severity: 'HIGH',
        context: { file_id: requestData.file_id }
      });
      
      return this.createErrorResponse(error.message, 500);
    }
  }

  /**
   * ヘルスチェックリクエストを処理します
   * @param {Object} requestData リクエストデータ
   * @param {string} apiKey APIキー（オプション）
   * @return {Object} APIレスポンス
   */
  static handleHealthCheckRequest(requestData, apiKey) {
    try {
      // システムステータスをチェック
      const config = Config.getSystemConfig();
      const status = {
        status: 'OK',
        version: config.version || '1.0.0',
        timestamp: new Date().toISOString(),
        components: {
          database: this.checkDatabaseConnection(),
          cache: this.checkCacheService(),
          search: true,
          response: true
        }
      };
      
      // 詳細情報のリクエスト（管理者のみ）
      if (requestData.detailed === 'true' && this.isAdminUser(apiKey)) {
        status.details = {
          database_id: config.database_id,
          chunk_count: this.getChunkCount(),
          document_count: this.getDocumentCount(),
          template_count: this.getTemplateCount(),
          memory_usage: this.getMemoryUsage()
        };
      }
      
      return this.createSuccessResponse(status);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.handleHealthCheckRequest',
        error: error,
        severity: 'LOW'
      });
      
      return this.createSuccessResponse({
        status: 'ERROR',
        message: error.message,
        timestamp: new Date().toISOString()
      });
    }
  }

  /**
   * データベース接続をチェックします
   * @return {boolean} 接続が正常かどうか
   */
  static checkDatabaseConnection() {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const sheets = ss.getSheets();
      return sheets.length > 0;
    } catch (error) {
      return false;
    }
  }

  /**
   * キャッシュサービスをチェックします
   * @return {boolean} キャッシュサービスが正常かどうか
   */
  static checkCacheService() {
    try {
      const testKey = 'health_check_' + Utilities.generateUniqueId();
      const testValue = 'test_value';
      
      CacheManager.set(testKey, testValue, 10);
      const retrievedValue = CacheManager.get(testKey);
      
      return retrievedValue === testValue;
    } catch (error) {
      return false;
    }
  }

  /**
   * チャンク数を取得します
   * @return {number} チャンク数
   */
  static getChunkCount() {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const indexSheet = ss.getSheetByName('Index_Mapping');
      
      if (!indexSheet) {
        return 0;
      }
      
      const data = indexSheet.getDataRange().getValues();
      let totalCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === 'chunks') {
          totalCount += parseInt(data[i][3]) || 0;
        }
      }
      
      return totalCount;
    } catch (error) {
      return 0;
    }
  }

  /**
   * ドキュメント数を取得します
   * @return {number} ドキュメント数
   */
  static getDocumentCount() {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const documentsSheet = ss.getSheetByName('Documents');
      
      if (!documentsSheet) {
        return 0;
      }
      
      return Math.max(0, documentsSheet.getLastRow() - 1); // ヘッダーを除く
    } catch (error) {
      return 0;
    }
  }

  /**
   * テンプレート数を取得します
   * @return {number} テンプレート数
   */
  static getTemplateCount() {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const templatesSheet = ss.getSheetByName('Templates');
      
      if (!templatesSheet) {
        return 0;
      }
      
      return Math.max(0, templatesSheet.getLastRow() - 1); // ヘッダーを除く
    } catch (error) {
      return 0;
    }
  }

  /**
   * メモリ使用量を取得します
   * @return {Object} メモリ使用状況
   */
  static getMemoryUsage() {
    // GASではメモリ使用量を直接取得できないため、推定値を返す
    return {
      estimated: true,
      cache_items: this.getApproximateCacheItemCount(),
      script_properties: this.getScriptPropertiesCount(),
      note: 'Memory usage information is approximate'
    };
  }

  /**
   * おおよそのキャッシュアイテム数を取得します
   * @return {number} キャッシュアイテム数
   */
  static getApproximateCacheItemCount() {
    try {
      // キャッシュから実際にはカウントできないため、推定値を返す
      return 'Not available in GAS environment';
    } catch (error) {
      return 0;
    }
  }

  /**
   * スクリプトプロパティの数を取得します
   * @return {number} プロパティの数
   */
  static getScriptPropertiesCount() {
    try {
      const props = PropertiesService.getScriptProperties();
      const keys = props.getKeys();
      return keys.length;
    } catch (error) {
      return 0;
    }
  }

  /**
   * 成功レスポンスを作成します
   * @param {Object} data レスポンスデータ
   * @return {Object} APIレスポンス
   */
  static createSuccessResponse(data) {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: data,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  /**
   * エラーレスポンスを作成します
   * @param {string} message エラーメッセージ
   * @param {number} code HTTPステータスコード
   * @return {Object} APIレスポンス
   */
  static createErrorResponse(message, code = 500) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: {
        message: message,
        code: code
      },
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  /**
   * 検索結果を最適化します
   * @param {Object} searchResults 検索結果
   * @return {Object} 最適化された検索結果
   */
  static optimizeSearchResults(searchResults) {
    try {
      if (!searchResults || !searchResults.success) {
        return searchResults;
      }
      
      // 結果をコピー
      const optimizedResults = { ...searchResults };
      
      // 検索結果のコンテンツを最適化
      if (optimizedResults.results && optimizedResults.results.length > 0) {
        optimizedResults.results = optimizedResults.results.map(result => {
          // 長いコンテンツはスニペットのみ保持
          if (result.content && result.content.length > 500) {
            return {
              ...result,
              content: result.snippet || result.content.substring(0, 500) + '...',
              content_truncated: true
            };
          }
          return result;
        });
      }
      
      // 検索結果のIDを生成して保存
      const resultId = Utilities.generateUniqueId();
      optimizedResults.result_id = resultId;
      
      // キャッシュに元のレスポンスを保存（30分）
      CacheManager.set(`search_results_${resultId}`, JSON.stringify(searchResults), 1800);
      
      return optimizedResults;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.optimizeSearchResults',
        error: error,
        severity: 'LOW'
      });
      
      return searchResults;
    }
  }

  /**
   * ドキュメントの埋め込み生成をスケジュールします
   * @param {string} documentId ドキュメントID
   */
  static scheduleEmbeddingGeneration(documentId) {
    try {
      // トリガーを作成して後で実行
      const trigger = ScriptApp.newTrigger('generateEmbeddingsForDocument')
        .timeBased()
        .after(1000) // 1秒後
        .create();
      
      // ドキュメントIDをプロパティに保存
      PropertiesService.getScriptProperties().setProperty(
        `embedding_trigger_${trigger.getUniqueId()}`,
        documentId
      );
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.scheduleEmbeddingGeneration',
        error: error,
        severity: 'MEDIUM',
        context: { document_id: documentId }
      });
    }
  }

  /**
   * ドキュメントの埋め込みを生成します（トリガーから呼び出される）
   * @param {Object} e トリガーイベント
   */
  static async generateEmbeddingsForDocument(e) {
    try {
      // トリガーIDからドキュメントIDを取得
      const triggerId = e.triggerUid;
      const documentId = PropertiesService.getScriptProperties().getProperty(`embedding_trigger_${triggerId}`);
      
      if (!documentId) {
        throw new Error('Document ID not found for trigger');
      }
      
      // プロパティを削除
      PropertiesService.getScriptProperties().deleteProperty(`embedding_trigger_${triggerId}`);
      
      // 埋め込みを生成
      await EmbeddingManager.generateEmbeddingsForDocument(documentId);
      
      // 完了後にトリガーを削除
      ScriptApp.deleteTrigger(ScriptApp.getProjectTriggers().find(
        trigger => trigger.getUniqueId() === triggerId
      ));
    } catch (error) {
      ErrorHandler.handleError({
        source: 'WebApi.generateEmbeddingsForDocument',
        error: error,
        severity: 'HIGH',
        context: { trigger_id: e?.triggerUid }
      });
    }
  }

  /**
   * フロントエンド設定を取得します
   * @return {Object} フロントエンド設定
   */
  static getFrontendConfig() {
    const config = Config.getSystemConfig();
    
    return {
      version: config.version || '1.0.0',
      theme_color: config.theme_color || '#1a73e8',
      supported_languages: config.supported_languages || ['ja', 'en'],
      default_language: config.default_language || 'ja',
      max_search_results: config.max_search_results || 10,
      semantic_search_weight: config.semantic_search_weight || 0.7,
      keyword_search_weight: config.keyword_search_weight || 0.3,
      gemini_api: {
        client_config: GeminiIntegration.getClientConfiguration()
      },
      categories: SearchEngine.getCategories(),
      response_types: [
        {id: 'standard', name: 'Standard'},
        {id: 'email', name: 'Email'},
        {id: 'prep', name: 'PREP Format'},
        {id: 'detailed', name: 'Detailed'}
      ]
    };
  }

  /**
   * ユーザーが認証されているかチェックします
   * @param {Object} e イベントオブジェクト
   * @return {boolean} 認証されているかどうか
   */
  static isUserAuthenticated(e) {
    try {
      // GASウェブアプリでは基本的にGoogleアカウントで認証される
      const userEmail = Session.getActiveUser().getEmail();
      
      // メールアドレスが空でないことを確認
      return userEmail && userEmail.length > 0;
    } catch (error) {
      return false;
    }
  }

  /**
   * ユーザーが管理者かどうかチェックします
   * @param {string} apiKey APIキー
   * @return {boolean} 管理者かどうか
   */
  static isAdminUser(apiKey) {
    try {
      // 現在のユーザーのメールアドレスを取得
      const userEmail = Session.getActiveUser().getEmail();
      
      // プロパティから管理者リストを取得
      const config = Config.getSystemConfig();
      const adminEmails = config.admin_emails || [];
      
      // APIキーが管理者のキーと一致するか、メールアドレスが管理者リストにあるかチェック
      return apiKey === Config.getAdminGeminiApiKey() || adminEmails.includes(userEmail);
    } catch (error) {
      return false;
    }
  }

  /**
   * スタイルコンテンツを取得します
   * @return {string} CSSコンテンツ
   */
  static getStylesContent() {
    // スタイルシートコンテンツを取得
    try {
      return HtmlService.createHtmlOutputFromFile('styles').getContent();
    } catch (error) {
      // スタイルファイルがない場合はデフォルトスタイルを返す
      return this.getDefaultStyles();
    }
  }

  /**
   * スクリプトコンテンツを取得します
   * @return {string} JavaScriptコンテンツ
   */
  static getScriptContent() {
    // スクリプトコンテンツを取得
    try {
      return HtmlService.createHtmlOutputFromFile('script').getContent();
    } catch (error) {
      // スクリプトファイルがない場合はデフォルトスクリプトを返す
      return this.getDefaultScript();
    }
  }

  /**
   * デフォルトのスタイルを取得します
   * @return {string} デフォルトCSS
   */
  static getDefaultStyles() {
    return `
      /* Material Design 3 Tokens - Light Theme */
      :root {
        --md-sys-color-primary: #1a73e8;
        --md-sys-color-on-primary: #ffffff;
        --md-sys-color-primary-container: #d3e3fd;
        --md-sys-color-on-primary-container: #0a1b38;
        --md-sys-color-secondary: #4285f4;
        --md-sys-color-on-secondary: #ffffff;
        --md-sys-color-secondary-container: #d8e6ff;
        --md-sys-color-on-secondary-container: #061441;
        --md-sys-color-tertiary: #34a853;
        --md-sys-color-on-tertiary: #ffffff;
        --md-sys-color-tertiary-container: #b7f0cc;
        --md-sys-color-on-tertiary-container: #05371a;
        --md-sys-color-error: #ea4335;
        --md-sys-color-on-error: #ffffff;
        --md-sys-color-error-container: #ffcdd2;
        --md-sys-color-on-error-container: #4c0f0b;
        --md-sys-color-background: #ffffff;
        --md-sys-color-on-background: #1f1f1f;
        --md-sys-color-surface: #ffffff;
        --md-sys-color-on-surface: #1f1f1f;
        --md-sys-color-surface-variant: #f8f9fa;
        --md-sys-color-on-surface-variant: #444746;
        --md-sys-color-outline: #dadce0;
        --md-sys-color-outline-variant: #f1f3f4;
        --md-sys-color-shadow: rgba(0,0,0,0.1);
        --md-sys-color-scrim: rgba(0,0,0,0.4);
        --md-sys-color-inverse-surface: #303134;
        --md-sys-color-inverse-on-surface: #f8f9fa;
        
        /* Typography */
        --md-sys-typescale-headline-large-font-family: 'Google Sans', Roboto, sans-serif;
        --md-sys-typescale-headline-large-font-size: 32px;
        --md-sys-typescale-headline-large-font-weight: 400;
        --md-sys-typescale-headline-large-line-height: 40px;
        
        --md-sys-typescale-title-large-font-family: 'Google Sans', Roboto, sans-serif;
        --md-sys-typescale-title-large-font-size: 22px;
        --md-sys-typescale-title-large-font-weight: 400;
        --md-sys-typescale-title-large-line-height: 28px;
        
        --md-sys-typescale-body-large-font-family: Roboto, sans-serif;
        --md-sys-typescale-body-large-font-size: 16px;
        --md-sys-typescale-body-large-font-weight: 400;
        --md-sys-typescale-body-large-line-height: 24px;
        
        --md-sys-typescale-body-medium-font-family: Roboto, sans-serif;
        --md-sys-typescale-body-medium-font-size: 14px;
        --md-sys-typescale-body-medium-font-weight: 400;
        --md-sys-typescale-body-medium-line-height: 20px;
        
        /* Shape */
        --md-sys-shape-corner-small: 4px;
        --md-sys-shape-corner-medium: 8px;
        --md-sys-shape-corner-large: 16px;
        --md-sys-shape-corner-extralarge: 28px;
        
        /* Elevation */
        --md-sys-elevation-level1: 0px 1px 2px rgba(0,0,0,0.1);
        --md-sys-elevation-level2: 0px 2px 4px rgba(0,0,0,0.1);
        --md-sys-elevation-level3: 0px 4px 8px rgba(0,0,0,0.1);
        --md-sys-elevation-level4: 0px 8px 16px rgba(0,0,0,0.1);
      }
      
      /* Material Design 3 Tokens - Dark Theme */
      @media (prefers-color-scheme: dark) {
        :root {
          --md-sys-color-primary: #8ab4f8;
          --md-sys-color-on-primary: #002a72;
          --md-sys-color-primary-container: #00419e;
          --md-sys-color-on-primary-container: #d3e3fd;
          --md-sys-color-secondary: #a8c7fa;
          --md-sys-color-on-secondary: #0d2e6b;
          --md-sys-color-secondary-container: #234591;
          --md-sys-color-on-secondary-container: #d8e6ff;
          --md-sys-color-tertiary: #8ada9e;
          --md-sys-color-on-tertiary: #0a6627;
          --md-sys-color-tertiary-container: #0e8735;
          --md-sys-color-on-tertiary-container: #b7f0cc;
          --md-sys-color-error: #f28b82;
          --md-sys-color-on-error: #690c05;
          --md-sys-color-error-container: #930f0a;
          --md-sys-color-on-error-container: #ffcdd2;
          --md-sys-color-background: #1f1f1f;
          --md-sys-color-on-background: #e3e3e3;
          --md-sys-color-surface: #1f1f1f;
          --md-sys-color-on-surface: #e3e3e3;
          --md-sys-color-surface-variant: #303134;
          --md-sys-color-on-surface-variant: #c4c7c5;
          --md-sys-color-outline: #747775;
          --md-sys-color-outline-variant: #444746;
          --md-sys-color-shadow: rgba(0,0,0,0.3);
          --md-sys-color-scrim: rgba(0,0,0,0.6);
          --md-sys-color-inverse-surface: #e3e3e3;
          --md-sys-color-inverse-on-surface: #303134;
        }
      }
      
      /* Base Styles */
      * {
        box-sizing: border-box;
        margin: 0;
        padding: 0;
      }
      
      body {
        font-family: var(--md-sys-typescale-body-large-font-family);
        font-size: var(--md-sys-typescale-body-large-font-size);
        line-height: var(--md-sys-typescale-body-large-line-height);
        color: var(--md-sys-color-on-background);
        background-color: var(--md-sys-color-background);
      }
      
      /* Layout */
      .container {
        max-width: 1280px;
        margin: 0 auto;
        padding: 16px;
      }
      
      /* Typography */
      h1 {
        font-family: var(--md-sys-typescale-headline-large-font-family);
        font-size: var(--md-sys-typescale-headline-large-font-size);
        font-weight: var(--md-sys-typescale-headline-large-font-weight);
        line-height: var(--md-sys-typescale-headline-large-line-height);
        margin-bottom: 24px;
        color: var(--md-sys-color-on-surface);
      }
      
      h2 {
        font-family: var(--md-sys-typescale-title-large-font-family);
        font-size: var(--md-sys-typescale-title-large-font-size);
        font-weight: var(--md-sys-typescale-title-large-font-weight);
        line-height: var(--md-sys-typescale-title-large-line-height);
        margin-bottom: 16px;
        color: var(--md-sys-color-on-surface);
      }
      
      /* Components */
      .card {
        background-color: var(--md-sys-color-surface);
        border-radius: var(--md-sys-shape-corner-medium);
        padding: 16px;
        margin-bottom: 16px;
        box-shadow: var(--md-sys-elevation-level1);
      }
      
      .button {
        background-color: var(--md-sys-color-primary);
        color: var(--md-sys-color-on-primary);
        border: none;
        border-radius: var(--md-sys-shape-corner-small);
        padding: 8px 16px;
        font-family: var(--md-sys-typescale-body-medium-font-family);
        font-size: var(--md-sys-typescale-body-medium-font-size);
        font-weight: 500;
        cursor: pointer;
        transition: background-color 0.2s ease;
      }
      
      .button:hover {
        background-color: var(--md-sys-color-primary-container);
        color: var(--md-sys-color-on-primary-container);
      }
      
      .text-field {
        display: flex;
        flex-direction: column;
        margin-bottom: 16px;
      }
      
      .text-field label {
        font-size: var(--md-sys-typescale-body-medium-font-size);
        margin-bottom: 8px;
        color: var(--md-sys-color-on-surface-variant);
      }
      
      .text-field input, .text-field textarea {
        font-family: var(--md-sys-typescale-body-large-font-family);
        font-size: var(--md-sys-typescale-body-large-font-size);
        padding: 12px;
        border: 1px solid var(--md-sys-color-outline);
        border-radius: var(--md-sys-shape-corner-small);
        background-color: var(--md-sys-color-surface);
        color: var(--md-sys-color-on-surface);
      }
      
      .text-field input:focus, .text-field textarea:focus {
        outline: 2px solid var(--md-sys-color-primary);
        border-color: transparent;
      }
      
      /* Responsive */
      @media (max-width: 600px) {
        .container {
          padding: 8px;
        }
        
        h1 {
          font-size: 24px;
          line-height: 32px;
        }
        
        h2 {
          font-size: 18px;
          line-height: 24px;
        }
      }
    `;
  }

  /**
   * デフォルトのスクリプトを取得します
   * @return {string} デフォルトJavaScript
   */
  static getDefaultScript() {
    return `
      // Configuration
      let config = {
        apiEndpoint: '',
        geminiApiKey: '',
        language: 'ja',
        responseType: 'standard'
      };
      
      // DOM Elements
      let elements = {};
      
      // Initialize the application
      async function initApp() {
        // Set API endpoint
        config.apiEndpoint = window.location.href.split('?')[0];
        
        // Get DOM elements
        elements = {
          searchForm: document.getElementById('search-form'),
          queryInput: document.getElementById('query-input'),
          apiKeyInput: document.getElementById('api-key-input'),
          resultsContainer: document.getElementById('results-container'),
          responseContainer: document.getElementById('response-container'),
          responseTypeSelect: document.getElementById('response-type-select'),
          languageSelect: document.getElementById('language-select'),
          loadingIndicator: document.getElementById('loading-indicator'),
          errorMessage: document.getElementById('error-message')
        };
        
        // Load configuration
        await loadConfig();
        
        // Initialize event listeners
        initEventListeners();
      }
      
      // Load application configuration
      async function loadConfig() {
        try {
          const response = await fetch(config.apiEndpoint + '?view=config');
          if (!response.ok) throw new Error('Failed to load configuration');
          
          const data = await response.json();
          
          // Update configuration
          config = { ...config, ...data };
          
          // Initialize UI with configuration
          initUiWithConfig();
        } catch (error) {
          console.error('Error loading configuration:', error);
          // Continue with default config
        }
      }
      
      // Initialize UI with configuration
      function initUiWithConfig() {
        // Set default language
        if (elements.languageSelect) {
          elements.languageSelect.value = config.default_language || 'ja';
        }
        
        // Populate language options
        if (elements.languageSelect && config.supported_languages) {
          elements.languageSelect.innerHTML = '';
          config.supported_languages.forEach(lang => {
            const option = document.createElement('option');
            option.value = lang;
            option.textContent = lang === 'ja' ? '日本語' : 'English';
            elements.languageSelect.appendChild(option);
          });
          elements.languageSelect.value = config.default_language || 'ja';
        }
        
        // Populate response type options
        if (elements.responseTypeSelect && config.response_types) {
          elements.responseTypeSelect.innerHTML = '';
          config.response_types.forEach(type => {
            const option = document.createElement('option');
            option.value = type.id;
            option.textContent = type.name;
            elements.responseTypeSelect.appendChild(option);
          });
        }
        
        // Load API key from local storage
        const savedApiKey = localStorage.getItem('gemini_api_key');
        if (savedApiKey && elements.apiKeyInput) {
          elements.apiKeyInput.value = savedApiKey;
        }
      }
      
      // Initialize event listeners
      function initEventListeners() {
        // Search form submission
        if (elements.searchForm) {
          elements.searchForm.addEventListener('submit', handleSearch);
        }
        
        // API key input change
        if (elements.apiKeyInput) {
          elements.apiKeyInput.addEventListener('change', () => {
            localStorage.setItem('gemini_api_key', elements.apiKeyInput.value);
          });
        }
        
        // Language selection change
        if (elements.languageSelect) {
          elements.languageSelect.addEventListener('change', () => {
            config.language = elements.languageSelect.value;
          });
        }
        
        // Response type selection change
        if (elements.responseTypeSelect) {
          elements.responseTypeSelect.addEventListener('change', () => {
            config.responseType = elements.responseTypeSelect.value;
          });
        }
      }
      
      // Handle search form submission
      async function handleSearch(event) {
        event.preventDefault();
        
        const query = elements.queryInput.value.trim();
        if (!query) return;
        
        // Show loading indicator
        showLoading(true);
        hideError();
        
        try {
          // Save API key to configuration
          config.geminiApiKey = elements.apiKeyInput.value;
          
          // Search request
          const searchResults = await performSearch(query);
          
          // Display search results
          displaySearchResults(searchResults);
          
          // Generate response from search results
          const response = await generateResponse(query, searchResults);
          
          // Display response
          displayResponse(response);
          
          // Hide loading indicator
          showLoading(false);
        } catch (error) {
          console.error('Search error:', error);
          showError(error.message || 'An error occurred during search');
          showLoading(false);
        }
      }
      
      // Perform search request
      async function performSearch(query) {
        const searchParams = new URLSearchParams({
          type: 'search',
          query: query,
          language: config.language,
          api_key: config.geminiApiKey
        });
        
        const response = await fetch(config.apiEndpoint, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
          },
          body: searchParams.toString()
        });
        
        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error?.message || 'Search request failed');
        }
        
        const data = await response.json();
        
        if (!data.success) {
          throw new Error(data.error?.message || 'Search failed');
        }
        
        return data.data;
      }
      
      // Generate response from search results
      async function generateResponse(query, searchResults) {
        if (!searchResults || !searchResults.result_id) return null;
        
        const responseParams = new URLSearchParams({
          type: 'generate_response',
          query: query,
          search_results_id: searchResults.result_id,
          response_type: config.responseType,
          language: config.language,
          api_key: config.geminiApiKey
        });
        
        const response = await fetch(config.apiEndpoint, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
          },
          body: responseParams.toString()
        });
        
        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error?.message || 'Response generation failed');
        }
        
        const data = await response.json();
        
        if (!data.success) {
          throw new Error(data.error?.message || 'Response generation failed');
        }
        
        return data.data;
      }
      
      // Display search results
      function displaySearchResults(searchResults) {
        if (!elements.resultsContainer) return;
        
        // Clear previous results
        elements.resultsContainer.innerHTML = '';
        
        if (!searchResults || !searchResults.results || searchResults.results.length === 0) {
          elements.resultsContainer.innerHTML = '<div class="card"><p>No results found</p></div>';
          return;
        }
        
        // Create results header
        const header = document.createElement('div');
        header.className = 'results-header';
        header.innerHTML = \`<h2>Search Results: \${searchResults.results.length} found</h2>\`;
        elements.resultsContainer.appendChild(header);
        
        // Create results list
        const resultsList = document.createElement('div');
        resultsList.className = 'results-list';
        
        // Add each result
        searchResults.results.forEach((result, index) => {
          const resultCard = document.createElement('div');
          resultCard.className = 'card result-card';
          
          // Create result content
          resultCard.innerHTML = \`
            <h3>\${result.title || 'Untitled Document'}</h3>
            <p class="result-category">\${result.category || 'General'}</p>
            <div class="result-snippet">
              <p>\${result.snippet || result.content || 'No preview available'}</p>
            </div>
            <div class="result-meta">
              <span class="result-relevance">Relevance: \${(result.relevance_score * 100).toFixed(0)}%</span>
              <span class="result-language">\${result.language || 'Unknown language'}</span>
            </div>
          \`;
          
          resultsList.appendChild(resultCard);
        });
        
        elements.resultsContainer.appendChild(resultsList);
      }
      
      // Display generated response
      function displayResponse(response) {
        if (!elements.responseContainer) return;
        
        // Clear previous response
        elements.responseContainer.innerHTML = '';
        
        if (!response || !response.content) {
          return;
        }
        
        // Create response card
        const responseCard = document.createElement('div');
        responseCard.className = 'card response-card';
        
        // Convert markdown to HTML
        const responseHtml = convertMarkdownToHtml(response.content);
        
        // Create response content
        responseCard.innerHTML = \`
          <h2>Response (\${response.response_type || 'standard'})</h2>
          <div class="response-content">\${responseHtml}</div>
          <div class="response-meta">
            <span class="response-language">\${response.language || 'Unknown language'}</span>
            <span class="response-time">Generated in \${response.generation_time_ms || 0}ms</span>
          </div>
        \`;
        
        elements.responseContainer.appendChild(responseCard);
      }
      
      // Convert markdown to HTML (simple version)
      function convertMarkdownToHtml(markdown) {
        if (!markdown) return '';
        
        // Replace headers
        let html = markdown
          .replace(/^### (.*$)/gim, '<h3>$1</h3>')
          .replace(/^## (.*$)/gim, '<h2>$1</h2>')
          .replace(/^# (.*$)/gim, '<h1>$1</h1>');
        
        // Replace emphasis
        html = html
          .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
          .replace(/\*(.*?)\*/g, '<em>$1</em>');
        
        // Replace lists
        html = html
          .replace(/^\s*- (.*$)/gim, '<ul><li>$1</li></ul>')
          .replace(/^\s*\d+\. (.*$)/gim, '<ol><li>$1</li></ol>');
        
        // Fix lists (merge adjacent lists)
        html = html
          .replace(/<\/ul>\s*<ul>/g, '')
          .replace(/<\/ol>\s*<ol>/g, '');
        
        // Replace blockquotes
        html = html
          .replace(/^\> (.*$)/gim, '<blockquote>$1</blockquote>');
        
        // Replace paragraphs (simple)
        html = html
          .replace(/^(?!<[a-z])(.*$)/gim, '<p>$1</p>');
        
        // Fix empty paragraphs
        html = html
          .replace(/<p><\/p>/g, '');
        
        return html;
      }
      
      // Show loading indicator
      function showLoading(isLoading) {
        if (!elements.loadingIndicator) return;
        
        elements.loadingIndicator.style.display = isLoading ? 'block' : 'none';
      }
      
      // Show error message
      function showError(message) {
        if (!elements.errorMessage) return;
        
        elements.errorMessage.textContent = message;
        elements.errorMessage.style.display = 'block';
      }
      
      // Hide error message
      function hideError() {
        if (!elements.errorMessage) return;
        
        elements.errorMessage.style.display = 'none';
      }
      
      // Initialize the application when the document is loaded
      document.addEventListener('DOMContentLoaded', initApp);
    `;
  }
}
/**
 * GASのWebアプリケーションエントリーポイント - GETリクエスト用
 * @param {Object} e イベントオブジェクト
 * @return {HtmlOutput} HTMLレスポンス
 */
function doGet(e) {
  return WebApi.doGet(e);
}

/**
 * GASのWebアプリケーションエントリーポイント - POSTリクエスト用
 * @param {Object} e イベントオブジェクト
 * @return {Object} APIレスポンス
 */
function doPost(e) {
  return WebApi.doPost(e);
}
