/**
 * RAG 2.0 多言語処理モジュール
 * 言語検出、前処理、翻訳機能など多言語対応処理を管理します
 */
class MultilangProcessor {

  /**
   * テキストの言語を検出します
   * @param {string} text 言語を検出するテキスト
   * @return {Object} 検出結果 {success: boolean, language: string, confidence: number}
   */
  static detectLanguage(text) {
    try {
      if (!text || text.trim().length === 0) {
        return {
          success: false,
          error: 'テキストが空です'
        };
      }

      // LanguageAppによる検出を試行
      let language;
      let confidence = 0.7; // デフォルトの信頼度
      
      try {
        language = LanguageApp.detect(text);
        confidence = 0.9; // LanguageAppによる検出は高信頼度
      } catch (e) {
        // LanguageAppが使えない場合はヒューリスティックな検出
        language = this.heuristicLanguageDetection(text);
        confidence = 0.7; // ヒューリスティック検出は中程度の信頼度
      }

      // サポートされていない言語の場合はデフォルト言語（英語）
      const supportedLanguages = Config.getSystemConfig().supported_languages;
      if (!supportedLanguages.includes(language)) {
        language = Config.getSystemConfig().default_language || 'en';
        confidence = 0.5; // デフォルト言語への置き換えは低信頼度
      }

      return {
        success: true,
        language: language,
        confidence: confidence
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'MultilangProcessor.detectLanguage',
        error: error,
        severity: 'LOW',
        context: { text_length: text ? text.length : 0 }
      });

      // エラー時はデフォルト言語（英語）を返す
      return {
        success: false,
        language: 'en',
        confidence: 0.3,
        error: error.message
      };
    }
  }

  /**
   * ヒューリスティックな言語検出を行います（LanguageAppがない場合のフォールバック）
   * @param {string} text 検出するテキスト
   * @return {string} 検出された言語コード
   */
  static heuristicLanguageDetection(text) {
    // 日本語の文字（ひらがな、カタカナ、漢字）を含むかチェック
    const jaCharRegex = /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/;
    if (jaCharRegex.test(text)) {
      return 'ja';
    }
    
    // 韓国語の文字（ハングル）を含むかチェック
    const koCharRegex = /[\uAC00-\uD7AF]/;
    if (koCharRegex.test(text)) {
      return 'ko';
    }
    
    // 中国語の文字（簡体字と繁体字）を含むかチェック
    const zhCharRegex = /[\u4E00-\u9FFF\u3400-\u4DBF]/;
    if (zhCharRegex.test(text) && !jaCharRegex.test(text)) {
      return 'zh';
    }
    
    // その他のケースは英語と仮定
    return 'en';
  }

  /**
   * テキストを言語に応じて前処理します
   * @param {string} text 前処理するテキスト
   * @param {string} language 言語コード
   * @return {Object} 前処理結果 {success: boolean, text: string, language: string}
   */
  static preprocessText(text, language = null) {
    try {
      if (!text || text.trim().length === 0) {
        return {
          success: false,
          error: 'テキストが空です'
        };
      }

      // 言語が指定されていない場合は検出
      if (!language) {
        const detection = this.detectLanguage(text);
        language = detection.language;
      }

      // 言語に応じた前処理
      let processedText;
      
      switch (language) {
        case 'ja':
          processedText = this.preprocessJapaneseText(text);
          break;
        case 'en':
        default:
          processedText = this.preprocessEnglishText(text);
          break;
      }

      return {
        success: true,
        text: processedText,
        language: language
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'MultilangProcessor.preprocessText',
        error: error,
        severity: 'LOW',
        context: { 
          text_length: text ? text.length : 0, 
          language: language 
        }
      });

      // エラー時は元のテキストを返す
      return {
        success: false,
        text: text,
        language: language || 'unknown',
        error: error.message
      };
    }
  }

  /**
   * 日本語テキストを前処理します
   * @param {string} text 前処理する日本語テキスト
   * @return {string} 前処理されたテキスト
   */
  static preprocessJapaneseText(text) {
    // Unicode正規化（NFC形式）
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
   * @param {string} text 前処理する英語テキスト
   * @param {boolean} removeStopwords ストップワードを除去するかどうか
   * @return {string} 前処理されたテキスト
   */
  static preprocessEnglishText(text, removeStopwords = false) {
    // Unicode正規化
    text = text.normalize('NFKC');
    
    // 小文字化
    text = text.toLowerCase();
    
    // 句読点の処理
    text = text.replace(/[^\w\s]/g, ' ');
    
    // 余分な空白の削除
    text = text.replace(/\s+/g, ' ').trim();
    
    // ストップワード除去（オプション）
    if (removeStopwords) {
      const stopwords = ['a', 'an', 'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'with'];
      const words = text.split(' ');
      return words.filter(word => !stopwords.includes(word)).join(' ');
    }
    
    return text;
  }

  /**
   * テキストをトークン化します
   * @param {string} text トークン化するテキスト
   * @param {string} language 言語コード
   * @return {Object} トークン化結果 {success: boolean, tokens: Array<string>, language: string}
   */
  static tokenizeText(text, language = null) {
    try {
      if (!text || text.trim().length === 0) {
        return {
          success: false,
          tokens: [],
          error: 'テキストが空です'
        };
      }

      // 言語が指定されていない場合は検出
      if (!language) {
        const detection = this.detectLanguage(text);
        language = detection.language;
      }

      // 言語に応じたトークン化
      let tokens;
      
      switch (language) {
        case 'ja':
          tokens = this.tokenizeJapaneseText(text);
          break;
        case 'en':
        default:
          tokens = this.tokenizeEnglishText(text);
          break;
      }

      return {
        success: true,
        tokens: tokens,
        language: language
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'MultilangProcessor.tokenizeText',
        error: error,
        severity: 'LOW',
        context: { 
          text_length: text ? text.length : 0, 
          language: language 
        }
      });

      // エラー時は単純分割
      return {
        success: false,
        tokens: text ? text.split(/\s+/) : [],
        language: language || 'unknown',
        error: error.message
      };
    }
  }

  /**
   * 日本語テキストをトークン化します
   * @param {string} text トークン化する日本語テキスト
   * @return {Array<string>} トークンの配列
   */
  static tokenizeJapaneseText(text) {
    // 日本語の場合、単純な文字単位の分割として実装
    // 注: 本番環境では形態素解析などより高度な手法を使用することを推奨
    
    // スペースで区切られた単語は分割を維持
    const segments = text.split(/\s+/);
    const tokens = [];
    
    for (const segment of segments) {
      if (segment.length === 0) continue;
      
      // 記号や句読点で分割
      const subSegments = segment.split(/([、。！？・「」『』（）［］【】〔〕…―]+)/);
      
      for (const subSegment of subSegments) {
        if (subSegment.length === 0) continue;
        
        // 記号や句読点以外の場合、さらに文字単位で処理
        if (!/^[、。！？・「」『』（）［］【】〔〕…―]+$/.test(subSegment)) {
          // 英数字の連続はまとめて1トークンに
          const charTokens = this.splitMixedJapaneseText(subSegment);
          tokens.push(...charTokens);
        } else {
          // 記号や句読点はそのままトークンとして追加
          tokens.push(subSegment);
        }
      }
    }
    
    return tokens;
  }

  /**
   * 日本語と英数字が混在するテキストを分割します
   * @param {string} text 分割するテキスト
   * @return {Array<string>} 分割結果
   */
  static splitMixedJapaneseText(text) {
    const result = [];
    let currentType = null;
    let currentToken = '';
    
    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      
      // 文字種の判定
      let type;
      if (/[a-zA-Z0-9]/.test(char)) {
        type = 'alnum'; // 英数字
      } else if (/[\u3040-\u309F]/.test(char)) {
        type = 'hiragana'; // ひらがな
      } else if (/[\u30A0-\u30FF]/.test(char)) {
        type = 'katakana'; // カタカナ
      } else if (/[\u4E00-\u9FAF]/.test(char)) {
        type = 'kanji'; // 漢字
      } else {
        type = 'other'; // その他
      }
      
      // 文字種の変わり目でトークンを区切る
      if (currentType === null) {
        // 初回
        currentType = type;
        currentToken = char;
      } else if (
        // 英数字とその他の文字種の切り替わり
        (currentType === 'alnum' && type !== 'alnum') ||
        (currentType !== 'alnum' && type === 'alnum') ||
        // その他の文字
        (currentType === 'other' || type === 'other')
      ) {
        // 文字種が変わったら現在のトークンを追加して新しいトークンを開始
        if (currentToken) {
          result.push(currentToken);
        }
        currentType = type;
        currentToken = char;
      } else {
        // 文字種が同じなら現在のトークンに追加
        currentToken += char;
      }
    }
    
    // 最後のトークンを追加
    if (currentToken) {
      result.push(currentToken);
    }
    
    return result;
  }

  /**
   * 英語テキストをトークン化します
   * @param {string} text トークン化する英語テキスト
   * @return {Array<string>} トークンの配列
   */
  static tokenizeEnglishText(text) {
    // 単語単位で分割
    return text.split(/\s+/).filter(token => token.length > 0);
  }

  /**
   * テキストのトークン数を推定します
   * @param {string} text トークン数を推定するテキスト
   * @param {string} language 言語コード
   * @return {number} 推定トークン数
   */
  static estimateTokenCount(text, language = null) {
    try {
      if (!text || text.trim().length === 0) {
        return 0;
      }

      // 言語が指定されていない場合は検出
      if (!language) {
        const detection = this.detectLanguage(text);
        language = detection.language;
      }

      // 言語に応じたトークン数推定
      let tokenCount;
      
      switch (language) {
        case 'ja':
          // 日本語の場合、文字数の0.5倍程度を目安に
          tokenCount = Math.ceil(text.length * 0.5);
          break;
        case 'en':
        default:
          // 英語の場合、単語数の1.3倍程度を目安に
          const words = text.split(/\s+/).filter(w => w.length > 0);
          tokenCount = Math.ceil(words.length * 1.3);
          break;
      }

      return tokenCount;
    } catch (error) {
      // エラー時は文字数をそのまま返す
      return text ? text.length : 0;
    }
  }

  /**
   * テキストを翻訳します
   * @param {string} text 翻訳するテキスト
   * @param {string} sourceLanguage 元の言語コード（自動検出する場合は null）
   * @param {string} targetLanguage 翻訳先の言語コード
   * @return {Object} 翻訳結果 {success: boolean, text: string, source_language: string, target_language: string}
   */
  static async translateText(text, sourceLanguage = null, targetLanguage = 'en') {
    try {
      if (!text || text.trim().length === 0) {
        return {
          success: false,
          error: 'テキストが空です'
        };
      }

      // 元の言語が指定されていない場合は検出
      if (!sourceLanguage) {
        const detection = this.detectLanguage(text);
        sourceLanguage = detection.language;
      }

      // 翻訳先言語の検証
      const supportedLanguages = Config.getSystemConfig().supported_languages;
      if (!supportedLanguages.includes(targetLanguage)) {
        targetLanguage = Config.getSystemConfig().default_language || 'en';
      }

      // 元の言語と翻訳先言語が同じ場合は翻訳不要
      if (sourceLanguage === targetLanguage) {
        return {
          success: true,
          text: text,
          source_language: sourceLanguage,
          target_language: targetLanguage,
          is_original: true
        };
      }

      // キャッシュをチェック
      const cacheKey = `translate_${sourceLanguage}_${targetLanguage}_${Utilities.generateUniqueId().substring(0, 8)}_${text.substring(0, 50)}`;
      const cachedTranslation = CacheManager.get(cacheKey);
      
      if (cachedTranslation) {
        const translationResult = JSON.parse(cachedTranslation);
        return {
          ...translationResult,
          from_cache: true
        };
      }

      // Gemini APIを使用した翻訳
      const translatedText = await this.translateWithGemini(text, sourceLanguage, targetLanguage);

      // 翻訳結果
      const result = {
        success: true,
        text: translatedText,
        source_language: sourceLanguage,
        target_language: targetLanguage,
        from_cache: false,
        is_original: false
      };

      // 翻訳結果をキャッシュに保存（1日）
      CacheManager.set(cacheKey, JSON.stringify(result), 86400);

      return result;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'MultilangProcessor.translateText',
        error: error,
        severity: 'MEDIUM',
        context: { 
          text_length: text ? text.length : 0, 
          source_language: sourceLanguage,
          target_language: targetLanguage
        }
      });

      // エラー時は元のテキストを返す
      return {
        success: false,
        text: text,
        source_language: sourceLanguage || 'unknown',
        target_language: targetLanguage,
        error: error.message,
        is_original: true
      };
    }
  }

  /**
   * Gemini APIを使用してテキストを翻訳します
   * @param {string} text 翻訳するテキスト
   * @param {string} sourceLanguage 元の言語コード
   * @param {string} targetLanguage 翻訳先の言語コード
   * @return {string} 翻訳されたテキスト
   */
  static async translateWithGemini(text, sourceLanguage, targetLanguage) {
    // 言語名の取得
    const sourceLanguageName = this.getLanguageName(sourceLanguage);
    const targetLanguageName = this.getLanguageName(targetLanguage);

    // 翻訳用プロンプトの作成
    const prompt = `Translate the following text from ${sourceLanguageName} to ${targetLanguageName}. Preserve the original formatting, including paragraphs, bullet points, and any special formatting. Provide only the translated text without any explanations:

${text}`;

    // Gemini APIで翻訳を実行
    const result = await GeminiIntegration.generateText(prompt, {
      temperature: 0.1, // 低い温度で正確な翻訳に
      maxOutputTokens: Math.max(this.estimateTokenCount(text, sourceLanguage) * 2, 1024),
      systemInstructions: [
        `You are a professional translator specializing in ${sourceLanguageName} to ${targetLanguageName} translation. Translate accurately while preserving the original meaning, tone, and formatting.`
      ]
    });

    if (!result.success || !result.text) {
      throw new Error(`翻訳に失敗しました: ${result.error || 'Unknown error'}`);
    }

    return result.text;
  }

  /**
   * 言語コードから言語名を取得します
   * @param {string} languageCode 言語コード
   * @return {string} 言語名
   */
  static getLanguageName(languageCode) {
    const languageNames = {
      'en': 'English',
      'ja': 'Japanese',
      'ko': 'Korean',
      'zh': 'Chinese',
      'fr': 'French',
      'de': 'German',
      'es': 'Spanish',
      'it': 'Italian',
      'pt': 'Portuguese',
      'ru': 'Russian'
    };

    return languageNames[languageCode] || 'Unknown';
  }

  /**
   * ヘルプページペアを取得します
   * @param {string} baseNameOrId ベース名またはID
   * @return {Object} ヘルプページペア情報
   */
  static getHelpPagePair(baseNameOrId) {
    try {
      // データベースを開く
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const helpPairsSheet = ss.getSheetByName('Help_Pairs');

      if (!helpPairsSheet) {
        throw new Error('Help_Pairs シートが見つかりません');
      }

      // シートからデータを取得
      const data = helpPairsSheet.getDataRange().getValues();
      if (data.length <= 1) {
        return null; // ヘッダー行のみまたは空のシート
      }

      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const pairIdIndex = headers.indexOf('pair_id');
      const baseNameIndex = headers.indexOf('base_name');
      const jaDocIdIndex = headers.indexOf('ja_document_id');
      const enDocIdIndex = headers.indexOf('en_document_id');
      const categoryIndex = headers.indexOf('category');
      const updatedAtIndex = headers.indexOf('updated_at');
      const statusIndex = headers.indexOf('status');

      // 必須フィールドの検証
      if (pairIdIndex === -1 || baseNameIndex === -1 || jaDocIdIndex === -1 || enDocIdIndex === -1) {
        return null;
      }

      // ベース名またはIDに一致する行を検索
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        if (row[pairIdIndex] === baseNameOrId || row[baseNameIndex] === baseNameOrId) {
          // ヘルプページペア情報を構築
          const pairInfo = {
            pair_id: row[pairIdIndex],
            base_name: row[baseNameIndex],
            ja_document_id: row[jaDocIdIndex],
            en_document_id: row[enDocIdIndex]
          };

          // オプションフィールドの追加
          if (categoryIndex !== -1) pairInfo.category = row[categoryIndex];
          if (updatedAtIndex !== -1) pairInfo.updated_at = row[updatedAtIndex];
          if (statusIndex !== -1) pairInfo.status = row[statusIndex];

          // 各言語のドキュメントを取得
          pairInfo.ja_document = SheetStorage.getDocumentById(pairInfo.ja_document_id);
          pairInfo.en_document = SheetStorage.getDocumentById(pairInfo.en_document_id);

          return pairInfo;
        }
      }

      return null; // 一致するペアが見つからない
    } catch (error) {
      ErrorHandler.handleError({
        source: 'MultilangProcessor.getHelpPagePair',
        error: error,
        severity: 'MEDIUM',
        context: { base_name_or_id: baseNameOrId }
      });

      return null;
    }
  }

  /**
   * 指定された言語のヘルプページを取得します
   * @param {string} baseNameOrId ベース名またはID
   * @param {string} language 言語コード ('ja' or 'en')
   * @return {Object} ヘルプページドキュメント情報
   */
  static getHelpPageByLanguage(baseNameOrId, language) {
    try {
      const pairInfo = this.getHelpPagePair(baseNameOrId);
      
      if (!pairInfo) {
        return null;
      }

      // 指定された言語のドキュメントを返す
      if (language === 'ja') {
        return pairInfo.ja_document;
      } else if (language === 'en') {
        return pairInfo.en_document;
      }

      // 言語が指定されていない場合はペア情報を返す
      return pairInfo;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'MultilangProcessor.getHelpPageByLanguage',
        error: error,
        severity: 'MEDIUM',
        context: { 
          base_name_or_id: baseNameOrId,
          language: language
        }
      });

      return null;
    }
  }

  /**
   * ドキュメントペアを取得します（ドキュメントIDから）
   * @param {string} documentId ドキュメントID
   * @return {Object} ドキュメントペア情報
   */
  static getDocumentPair(documentId) {
    try {
      // ドキュメントを取得
      const document = SheetStorage.getDocumentById(documentId);
      
      if (!document || !document.metadata) {
        return null;
      }

      // ペアのドキュメントIDを取得
      const pairedDocumentId = document.metadata.paired_document_id;
      
      if (!pairedDocumentId) {
        return null;
      }

      // ペアのドキュメントを取得
      const pairedDocument = SheetStorage.getDocumentById(pairedDocumentId);
      
      if (!pairedDocument) {
        return null;
      }

      // ペア情報を構築
      return {
        original_document: document,
        paired_document: pairedDocument,
        original_language: document.language || 'unknown',
        paired_language: pairedDocument.language || 'unknown'
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'MultilangProcessor.getDocumentPair',
        error: error,
        severity: 'MEDIUM',
        context: { document_id: documentId }
      });

      return null;
    }
  }

  /**
   * ドキュメントを指定された言語で取得します
   * @param {string} documentId ドキュメントID
   * @param {string} targetLanguage 目標言語コード
   * @return {Object} 指定言語のドキュメント
   */
  static getDocumentInLanguage(documentId, targetLanguage) {
    try {
      // ドキュメントを取得
      const document = SheetStorage.getDocumentById(documentId);
      
      if (!document) {
        return null;
      }

      // ドキュメントが既に目標言語の場合はそのまま返す
      if (document.language === targetLanguage) {
        return document;
      }

      // ペアを取得
      const pair = this.getDocumentPair(documentId);
      
      if (pair && pair.paired_document.language === targetLanguage) {
        return pair.paired_document;
      }

      // ペアが見つからないか目標言語でない場合はnullを返す
      return null;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'MultilangProcessor.getDocumentInLanguage',
        error: error,
        severity: 'MEDIUM',
        context: { 
          document_id: documentId,
          target_language: targetLanguage
        }
      });

      return null;
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
        case 'detectLanguage':
          const languageResult = this.detectLanguage(params.text);
          return languageResult.success;
          
        case 'preprocessText':
          const preprocessResult = this.preprocessText(params.text, params.language);
          return preprocessResult.success;
          
        case 'tokenizeText':
          const tokenizeResult = this.tokenizeText(params.text, params.language);
          return tokenizeResult.success;
          
        case 'translateText':
          const translateResult = await this.translateText(params.text, params.sourceLanguage, params.targetLanguage);
          return translateResult.success;
          
        default:
          console.warn(`Unknown operation for retry: ${operation}`);
          return false;
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'MultilangProcessor.retryOperation',
        error: error,
        severity: 'HIGH',
        context: options.context
      });
      
      return false;
    }
  }
}
