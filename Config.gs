/**
 * RAG 2.0 システム設定
 * システム全体で使用される設定を管理します
 */
class Config {

  /**
   * システム設定を取得します
   * スプレッドシートの Config シートから設定を読み込みます
   * @return {Object} 設定オブジェクト
   */
  static getSystemConfig() {
    try {
      // キャッシュから設定を取得
      const cachedConfig = CacheManager.get('system_config');
      if (cachedConfig) {
        return JSON.parse(cachedConfig);
      }

      // スプレッドシートから設定を読み込み
      const ss = SpreadsheetApp.openById(this.getDatabaseId());
      const configSheet = ss.getSheetByName('Config');

      if (!configSheet) {
        throw new Error('設定シートが見つかりません');
      }

      const data = configSheet.getDataRange().getValues();
      const config = {};

      // ヘッダー行をスキップして設定を読み込み
      for (let i = 1; i < data.length; i++) {
        const key = data[i][0];
        const value = data[i][1];
        const type = data[i][2];

        // 型に応じた変換
        if (type === 'number') {
          config[key] = Number(value);
        } else if (type === 'boolean') {
          config[key] = value.toLowerCase() === 'true';
        } else if (type === 'json') {
          try {
            config[key] = JSON.parse(value);
          } catch (e) {
            config[key] = value;
          }
        } else {
          config[key] = value;
        }
      }

      // キャッシュに設定を保存 (1時間)
      CacheManager.set('system_config', JSON.stringify(config), 3600);

      return config;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'Config.getSystemConfig',
        error: error,
        severity: 'HIGH'
      });

      // デフォルト設定を返す
      return this.getDefaultConfig();
    }
  }

  /**
   * デフォルト設定を返します
   * 設定読み込みに失敗した場合に使用されます
   * @return {Object} デフォルト設定
   */
  static getDefaultConfig() {
    return {
      // データベース設定
      database_id: PropertiesService.getScriptProperties().getProperty('DATABASE_ID') || '',
      max_rows_per_sheet: 38000, // シートあたりの最大行数 (40000の95%)
      chunk_size: 500, // チャンクサイズ
      chunk_overlap: 100, // チャンク重複

      // 埋め込み設定
      embedding_dimension: 512, // 埋め込みベクトルの次元数
      embedding_model: 'models/embedding-001', // Gemini埋め込みモデル

      // 生成設定
      generation_model: 'models/gemini-1.5-pro', // Gemini生成モデル

      // 検索設定
      semantic_search_weight: 0.7, // セマンティック検索の重み
      keyword_search_weight: 0.3, // キーワード検索の重み
      max_search_results: 10, // 最大検索結果数

      // キャッシュ設定
      cache_ttl_short: 300, // 短期キャッシュTTL (秒)
      cache_ttl_medium: 3600, // 中期キャッシュTTL (秒)
      cache_ttl_long: 86400, // 長期キャッシュTTL (秒)

      // エラー処理設定
      max_retry_count: 3, // 最大再試行回数
      retry_backoff_base: 2, // 再試行バックオフの基数

      // 言語設定
      supported_languages: ['ja', 'en'], // サポートされる言語
      default_language: 'ja', // デフォルト言語

      // API設定
      api_timeout: 30000, // API タイムアウト (ミリ秒)

      // メンテナンス設定
      daily_backup_hour: 1, // 日次バックアップ時刻 (0-23)
      weekly_backup_day: 0, // 週次バックアップ曜日 (0=日曜)

      // UI設定
      theme_color: '#1a73e8', // テーマカラー

      // バージョン情報
      version: '1.0.0'
    };
  }

  /**
   * データベースIDを取得します
   * @return {string} スプレッドシートID
   */
  static getDatabaseId() {
    // スクリプトプロパティから取得
    const databaseId = PropertiesService.getScriptProperties().getProperty('DATABASE_ID');

    if (!databaseId) {
      throw new Error('DATABASE_ID がスクリプトプロパティに設定されていません');
    }

    return databaseId;
  }

  /**
   * 管理者のGemini API Keyを取得します
   * @return {string} API Key
   */
  static getAdminGeminiApiKey() {
    // スクリプトプロパティから取得
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

    if (!apiKey) {
      throw new Error('GEMINI_API_KEY がスクリプトプロパティに設定されていません');
    }

    return apiKey;
  }

  /**
   * RAGドキュメントのルートフォルダIDを取得します
   * @return {string} フォルダID
   */
  static getRootFolderId() {
    // スクリプトプロパティから取得
    const folderId = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');

    if (!folderId) {
      throw new Error('ROOT_FOLDER_ID がスクリプトプロパティに設定されていません');
    }

    return folderId;
  }

  /**
   * カテゴリ別のフォルダIDマッピングを取得します
   * @return {Object} カテゴリとフォルダIDのマッピング
   */
  static getCategoryFolderMapping() {
    const config = this.getSystemConfig();
    return config.category_folder_mapping || {
      'Help_Pages': '',
      'Search': '',
      'Mobile': '',
      'Shopping': '',
      'Display': '',
      'Video': '',
      'M&A': '',
      'Billing': '',
      'Policy': ''
    };
  }

  /**
   * シートネーミングルールを取得します
   * @param {string} type シートタイプ ('chunks' or 'embeddings')
   * @param {string} category カテゴリ名
   * @param {number} shardNumber シャード番号
   * @return {string} シート名
   */
  static getSheetName(type, category, shardNumber) {
    if (type === 'chunks') {
      return `Chunks_${category}_${shardNumber}`;
    } else if (type === 'embeddings') {
      return `Embeddings_${category}_${shardNumber}`;
    }
    return '';
  }
}
