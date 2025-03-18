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

        // ヘッダー行を取
