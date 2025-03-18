/**
 * RAG 2.0 ユーティリティクラス
 * システム全体で使用される共通ユーティリティ関数を提供します
 */
class Utilities {

  /**
   * 一意のIDを生成します
   * @return {string} 一意のID
   */
  static generateUniqueId() {
    return Utilities.getUuid();
  }

  /**
   * UUIDを生成します
   * @return {string} UUID
   */
  static getUuid() {
    return Utilities.formatUuid(Utilities.computeUuid());
  }

  /**
   * UUIDを計算します
   * @return {string} 未フォーマットのUUID
   */
  static computeUuid() {
    const chars = '0123456789abcdef'.split('');
    const uuid = [];

    for (let i = 0; i < 36; i++) {
      if (i === 8 || i === 13 || i === 18 || i === 23) {
        uuid[i] = '-';
      } else if (i === 14) {
        uuid[i] = '4';
      } else if (i === 19) {
        uuid[i] = chars[(Math.random() * 4) | 8];
      } else {
        uuid[i] = chars[(Math.random() * 16) | 0];
      }
    }

    return uuid.join('');
  }

  /**
   * UUIDをフォーマットします
   * @param {string} uuid 未フォーマットのUUID
   * @return {string} フォーマット済みUUID
   */
  static formatUuid(uuid) {
    return uuid.replace(/-/g, '');
  }

  /**
   * バイト配列をBase64文字列に変換します
   * @param {Uint8Array} bytes バイト配列
   * @return {string} Base64文字列
   */
  static bytesToBase64(bytes) {
    return Utilities.base64Encode(bytes);
  }

  /**
   * Base64文字列をバイト配列に変換します
   * @param {string} base64 Base64文字列
   * @return {Uint8Array} バイト配列
   */
  static base64ToBytes(base64) {
    return Utilities.base64Decode(base64);
  }

  /**
   * テキストをGZIP圧縮してBase64エンコードします
   * 埋め込みベクトルの保存に使用します
   * @param {string} text 圧縮するテキスト
   * @return {string} 圧縮されたBase64文字列
   */
  static compressAndEncodeText(text) {
    const blob = Utilities.newBlob(text);
    const compressed = Utilities.gzip(blob);
    return Utilities.base64Encode(compressed.getBytes());
  }

  /**
   * Base64エンコードされたGZIP圧縮テキストを復元します
   * @param {string} compressedBase64 圧縮されたBase64文字列
   * @return {string} 復元されたテキスト
   */
  static decodeAndDecompressText(compressedBase64) {
    const bytes = Utilities.base64Decode(compressedBase64);
    const blob = Utilities.newBlob(bytes);
    const uncompressed = Utilities.ungzip(blob);
    return uncompressed.getDataAsString();
  }

  /**
   * 配列を圧縮してBase64エンコードします
   * 埋め込みベクトルの保存に使用します
   * @param {Array<number>} array 数値配列
   * @return {string} 圧縮されたBase64文字列
   */
  static compressAndEncodeArray(array) {
    // 配列を小数点第4位までに丸める（精度と容量のバランス）
    const roundedArray = array.map(value => Math.round(value * 10000) / 10000);
    
    // 配列をJSON文字列に変換
    const json = JSON.stringify(roundedArray);
    // 圧縮してBase64エンコード
    return this.compressAndEncodeText(json);
  }

  /**
   * Base64エンコードされた圧縮配列を復元します
   * @param {string} compressedBase64 圧縮されたBase64文字列
   * @return {Array<number>} 復元された数値配列
   */
  static decodeAndDecompressArray(compressedBase64) {
    // 圧縮テキストを復元
    const json = this.decodeAndDecompressText(compressedBase64);
    // JSON文字列を配列に変換
    return JSON.parse(json);
  }

  /**
   * テキストの言語を検出します
   * @param {string} text 言語を検出するテキスト
   * @return {string} 言語コード ('ja', 'en' など)
   */
  static detectLanguage(text) {
    if (!text || text.trim().length === 0) {
      return 'en'; // デフォルト言語
    }

    try {
      const language = LanguageApp.detect(text);
      // サポートされている言語に制限
      const supportedLanguages = Config.getSystemConfig().supported_languages;
      if (supportedLanguages.includes(language)) {
        return language;
      } else {
        return 'en'; // デフォルト言語
      }
    } catch (error) {
      // LanguageAppが使えない場合はヒューリスティックな判定
      const jaCharRegex = /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/;
      if (jaCharRegex.test(text)) {
        return 'ja';
      }
      
      ErrorHandler.handleError({
        source: 'Utilities.detectLanguage',
        error: error,
        severity: 'LOW'
      });
      return 'en'; // デフォルト言語
    }
  }

  /**
   * テキストをチャンクに分割します
   * @param {string} text 分割するテキスト
   * @param {number} chunkSize チャンクサイズ
   * @param {number} overlap チャンク間の重複
   * @return {Array<string>} チャンクの配列
   */
  static splitTextIntoChunks(text, chunkSize, overlap) {
    if (!text) return [];

    const config = Config.getSystemConfig();
    chunkSize = chunkSize || config.chunk_size;
    overlap = overlap || config.chunk_overlap;

    // テキストをパラグラフに分割
    const paragraphs = text.split(/\n\s*\n/);
    const chunks = [];
    let currentChunk = '';

    for (const paragraph of paragraphs) {
      // パラグラフを追加するとチャンクサイズを超える場合
      if (currentChunk.length + paragraph.length > chunkSize && currentChunk.length > 0) {
        // 現在のチャンクを保存
        chunks.push(currentChunk);
        // 新しいチャンクを開始（重複部分を含む）
        const lastWords = currentChunk.split(' ').slice(-overlap).join(' ');
        currentChunk = lastWords + ' ' + paragraph;
      } else {
        // パラグラフを現在のチャンクに追加
        if (currentChunk.length > 0) {
          currentChunk += '\n\n';
        }
        currentChunk += paragraph;
      }
    }

    // 最後のチャンクを追加
    if (currentChunk.length > 0) {
      chunks.push(currentChunk);
    }

    return chunks;
  }

  /**
   * テキストから基本的なメタデータを抽出します
   * @param {string} text 処理するテキスト
   * @param {Object} existingMetadata 既存のメタデータ
   * @return {Object} 抽出されたメタデータ
   */
  static extractMetadataFromText(text, existingMetadata = {}) {
    const metadata = {...existingMetadata};

    // タイトルの抽出 (Markdown形式のヘッダーを検索)
    if (!metadata.title) {
      const titleMatch = text.match(/^#\s+(.+)$/m) || text.match(/^(.+)\n={3,}$/m);
      if (titleMatch) {
        metadata.title = titleMatch[1].trim();
      } else {
        // 最初の行をタイトルとして使用
        const firstLine = text.split('\n')[0].trim();
        if (firstLine && firstLine.length < 100) { // 短すぎない場合
          metadata.title = firstLine;
        }
      }
    }

    // 言語の検出
    if (!metadata.language) {
      metadata.language = this.detectLanguage(text);
    }

    // キーワードの抽出
    if (!metadata.keywords) {
      const words = text.toLowerCase()
        .replace(/[^\w\s]/g, ' ')
        .split(/\s+/)
        .filter(word => word.length > 3);

      const wordCounts = {};
      for (const word of words) {
        wordCounts[word] = (wordCounts[word] || 0) + 1;
      }

      const sortedWords = Object.entries(wordCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(entry => entry[0]);

      metadata.keywords = sortedWords;
    }

    return metadata;
  }

  /**
   * 現在のタイムスタンプを取得します
   * @return {string} ISO形式のタイムスタンプ
   */
  static getCurrentTimestamp() {
    return new Date().toISOString();
  }

  /**
   * オブジェクトのディープコピーを作成します
   * @param {Object} obj コピーするオブジェクト
   * @return {Object} コピーされたオブジェクト
   */
  static deepCopy(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  /**
   * テキストをトークン化します
   * @param {string} text トークン化するテキスト
   * @param {string} language 言語コード
   * @return {Array<string>} トークンの配列
   */
  static tokenizeText(text, language) {
    if (!text) return [];

    language = language || this.detectLanguage(text);

    if (language === 'ja') {
      // 日本語の場合は文字単位で分割
      return text.replace(/\s+/g, ' ').split('');
    } else {
      // 英語などの場合は単語単位で分割
      return text.replace(/\s+/g, ' ').split(/\b/).filter(t => t.trim().length > 0);
    }
  }

  /**
   * テキストの概算トークン数を計算します
   * @param {string} text 計算するテキスト
   * @param {string} language 言語コード
   * @return {number} 概算トークン数
   */
  static estimateTokenCount(text, language) {
    if (!text) return 0;

    language = language || this.detectLanguage(text);

    // 日本語と英語のトークン数を概算
    if (language === 'ja') {
      // 日本語の場合：文字数 * 0.5 (Unicode文字は通常0.5〜2トークン)
      return Math.ceil(text.length * 0.5);
    } else {
      // 英語の場合：単語数 * 1.3 (平均的な係数)
      const words = text.split(/\s+/).filter(w => w.length > 0);
      return Math.ceil(words.length * 1.3);
    }
  }
  
  /**
   * 日本語テキストを前処理します
   * @param {string} text 処理するテキスト
   * @return {string} 前処理されたテキスト
   */
  static preprocessJapaneseText(text) {
    // Unicode正規化（NFC、NFKC、NFD、NFKDの違いに注意）
    text = text.normalize('NFKC');
    
    // 全角/半角の統一（数字・アルファベット・記号）
    text = text.replace(/[！-～]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });
    
    // 句読点の正規化
    text = text.replace(/[、，]/g, "、")
               .replace(/[。．]/g, "。");
    
    // 余分な空白の削除（必要に応じて調整）
    text = text.replace(/\s+/g, " ").trim();
    
    return text;
  }
  
  /**
   * 英語テキストを前処理します
   * @param {string} text 処理するテキスト
   * @param {boolean} removeStopwords ストップワードを除去するかどうか
   * @return {string} 前処理されたテキスト
   */
  static preprocessEnglishText(text, removeStopwords = false) {
    // Unicode正規化
    text = text.normalize('NFKC');
    
    // 小文字化（検索のため）
    text = text.toLowerCase();
    
    // 句読点の置換と空白の正規化
    text = text.replace(/[^\w\s]/g, ' ')
               .replace(/\s+/g, ' ')
               .trim();
    
    // ストップワード除去（オプション）
    if (removeStopwords) {
      const stopwords = ['a', 'an', 'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'with'];
      const words = text.split(' ');
      return words.filter(word => !stopwords.includes(word)).join(' ');
    }
    
    return text;
  }
}
