/**
 * RAG 2.0 Gemini API 統合
 * Gemini AI との安全で効率的な通信を管理します
 */
class GeminiIntegration {

  /**
   * 管理者の API キーを使用してテキストからエンベディングを生成します
   * @param {string} text エンベディングを生成するテキスト
   * @param {string} model エンベディングモデル (デフォルト: models/embedding-001)
   * @return {Object} 生成結果 {success: boolean, embedding: number[], error: string}
   */
  static async generateEmbedding(text, model = null) {
    try {
      if (!text || text.trim().length === 0) {
        throw new Error('エンベディングを生成するテキストが空です');
      }

      // モデルが指定されていない場合はデフォルト値を使用
      if (!model) {
        model = Config.getSystemConfig().embedding_model;
      }

      // API キーを取得
      const apiKey = Config.getAdminGeminiApiKey();
      if (!apiKey) {
        throw new Error('Gemini API キーが構成されていません');
      }

      // API リクエストデータを準備
      const requestData = {
        model: model,
        content: { parts: [{ text: text }] }
      };

      // API リクエストを実行
      const response = await this.makeApiRequest('embedContent', requestData, apiKey);

      // レスポンスを検証
      if (!response || !response.embedding || !response.embedding.values) {
        throw new Error('API からの応答が無効です');
      }

      return {
        success: true,
        embedding: response.embedding.values,
        model: model
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'GeminiIntegration.generateEmbedding',
        error: error,
        severity: 'MEDIUM',
        context: { text_length: text ? text.length : 0, model: model },
        retry: this.isRetryableError(error)
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * 管理者の API キーを使用してテキスト生成を行います
   * @param {string} prompt 生成プロンプト
   * @param {Object} options 生成オプション
   * @param {string} options.model モデル名
   * @param {number} options.temperature 温度 (0.0-1.0)
   * @param {number} options.maxOutputTokens 最大出力トークン数
   * @param {Array<Object>} options.systemInstructions システム指示
   * @return {Object} 生成結果 {success: boolean, text: string, error: string}
   */
  static async generateText(prompt, options = {}) {
    try {
      if (!prompt || prompt.trim().length === 0) {
        throw new Error('生成プロンプトが空です');
      }

      // デフォルトオプション
      const config = Config.getSystemConfig();
      const defaultOptions = {
        model: config.generation_model,
        temperature: 0.7,
        maxOutputTokens: 2048,
        systemInstructions: []
      };

      // オプションをマージ
      const mergedOptions = { ...defaultOptions, ...options };

      // API キーを取得
      const apiKey = Config.getAdminGeminiApiKey();
      if (!apiKey) {
        throw new Error('Gemini API キーが構成されていません');
      }

      // API リクエストデータを準備
      const requestData = {
        model: mergedOptions.model,
        contents: [
          {
            parts: [
              { text: prompt }
            ]
          }
        ],
        generationConfig: {
          temperature: mergedOptions.temperature,
          maxOutputTokens: mergedOptions.maxOutputTokens,
          topP: 0.9,
          topK: 40
        }
      };

      // システム指示がある場合は追加
      if (mergedOptions.systemInstructions && mergedOptions.systemInstructions.length > 0) {
        requestData.systemInstructions = {
          parts: mergedOptions.systemInstructions.map(instruction => {
            return { text: instruction };
          })
        };
      }

      // API リクエストを実行
      const response = await this.makeApiRequest('generateContent', requestData, apiKey);

      // レスポンスを検証
      if (!response || !response.candidates || response.candidates.length === 0) {
        throw new Error('API からの応答が無効です');
      }

      const generatedText = response.candidates[0].content.parts[0].text;

      return {
        success: true,
        text: generatedText,
        model: mergedOptions.model
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'GeminiIntegration.generateText',
        error: error,
        severity: 'MEDIUM',
        context: { 
          prompt_length: prompt ? prompt.length : 0,
          model: options.model
        },
        retry: this.isRetryableError(error)
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * クライアントサイドで使用するための設定を返します
   * （API キーは含まれません）
   * @return {Object} クライアント設定
   */
  static getClientConfiguration() {
    const config = Config.getSystemConfig();
    
    return {
      models: {
        embedding: config.embedding_model,
        generation: config.generation_model
      },
      defaultSettings: {
        temperature: 0.7,
        maxOutputTokens: 2048
      },
      apiEndpoints: {
        embedContent: 'https://generativelanguage.googleapis.com/v1beta/models/{model}:embedContent',
        generateContent: 'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent'
      }
    };
  }

  /**
   * API リクエストを実行します
   * @param {string} endpoint API エンドポイント ('embedContent' or 'generateContent')
   * @param {Object} data リクエストデータ
   * @param {string} apiKey API キー
   * @return {Object} API レスポンス
   */
  static async makeApiRequest(endpoint, data, apiKey) {
    // エンドポイント URL を作成
    let url;
    
    if (endpoint === 'embedContent') {
      url = `https://generativelanguage.googleapis.com/v1beta/${data.model}:embedContent?key=${apiKey}`;
    } else if (endpoint === 'generateContent') {
      url = `https://generativelanguage.googleapis.com/v1beta/${data.model}:generateContent?key=${apiKey}`;
    } else {
      throw new Error(`無効なエンドポイント: ${endpoint}`);
    }

    // ヘッダーを設定
    const headers = {
      'Content-Type': 'application/json'
    };

    // fetch オプションを設定
    const options = {
      method: 'POST',
      headers: headers,
      muteHttpExceptions: true,
      payload: JSON.stringify(data)
    };

    // API timeout 設定
    const timeout = Config.getSystemConfig().api_timeout || 30000;

    try {
      // リクエストを実行
      const startTime = new Date().getTime();
      const response = UrlFetchApp.fetch(url, options);
      const endTime = new Date().getTime();
      
      // レスポンスを解析
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      const responseJson = JSON.parse(responseText);
      
      // レスポンスコードをチェック
      if (responseCode >= 200 && responseCode < 300) {
        // 成功レスポンス
        return responseJson;
      } else {
        // エラーレスポンス
        const error = new Error(responseJson.error?.message || `API エラー: ${responseCode}`);
        error.code = responseCode;
        error.details = responseJson.error;
        throw error;
      }
    } catch (error) {
      // タイムアウトチェック
      if (new Date().getTime() - startTime >= timeout) {
        throw new Error(`API リクエストがタイムアウトしました: ${timeout}ms`);
      }
      
      throw error;
    }
  }

  /**
   * エラーが再試行可能かどうかを判断します
   * @param {Error} error エラーオブジェクト
   * @return {boolean} 再試行可能かどうか
   */
  static isRetryableError(error) {
    // エラーメッセージを取得
    const errorMessage = error.message || '';
    
    // 再試行可能なエラーパターン
    const retryablePatterns = [
      /timeout/i,
      /rate limit/i,
      /quota/i,
      /temporarily unavailable/i,
      /server error/i,
      /503/i,
      /500/i,
      /429/i
    ];
    
    // 再試行不可能なエラーパターン
    const nonRetryablePatterns = [
      /invalid api key/i,
      /authentication/i,
      /permission denied/i,
      /not found/i,
      /invalid request/i,
      /400/i,
      /401/i,
      /403/i,
      /404/i
    ];
    
    // 再試行不可能なエラーの場合は false
    if (nonRetryablePatterns.some(pattern => pattern.test(errorMessage))) {
      return false;
    }
    
    // 再試行可能なエラーの場合は true
    if (retryablePatterns.some(pattern => pattern.test(errorMessage))) {
      return true;
    }
    
    // デフォルトでは再試行可能とみなす
    return true;
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
        case 'generateEmbedding':
          const embeddingResult = await this.generateEmbedding(params.text, params.model);
          return embeddingResult.success;
          
        case 'generateText':
          const textResult = await this.generateText(params.prompt, params.options);
          return textResult.success;
          
        default:
          console.warn(`Unknown operation for retry: ${operation}`);
          return false;
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'GeminiIntegration.retryOperation',
        error: error,
        severity: 'HIGH',
        context: options.context
      });
      
      return false;
    }
  }

  /**
   * テキストのバッチに対してエンベディングを生成します
   * 大量のテキストを処理する場合に効率的
   * @param {Array<string>} texts エンベディングを生成するテキストの配列
   * @param {string} model エンベディングモデル (デフォルト: models/embedding-001)
   * @return {Object} 生成結果 {success: boolean, embeddings: Array<{text: string, embedding: number[]}>, errors: Array}
   */
  static async generateEmbeddingBatch(texts, model = null) {
    try {
      if (!texts || !Array.isArray(texts) || texts.length === 0) {
        throw new Error('エンベディングを生成するテキストの配列が無効です');
      }

      // モデルが指定されていない場合はデフォルト値を使用
      if (!model) {
        model = Config.getSystemConfig().embedding_model;
      }

      const results = [];
      const errors = [];
      
      // テキストをバッチ処理（API制限を考慮して小さなバッチに分割）
      const batchSize = 20; // 一度に処理するテキストの数
      for (let i = 0; i < texts.length; i += batchSize) {
        const batch = texts.slice(i, i + batchSize);
        
        // 並列処理ではなく逐次処理（API制限を考慮）
        for (const text of batch) {
          if (!text || text.trim().length === 0) {
            errors.push({
              text: text,
              error: 'テキストが空です'
            });
            continue;
          }
          
          // エンベディング生成
          const result = await this.generateEmbedding(text, model);
          
          if (result.success) {
            results.push({
              text: text,
              embedding: result.embedding
            });
          } else {
            errors.push({
              text: text,
              error: result.error
            });
          }
          
          // API制限を考慮して少し待機
          Utilities.sleep(500);
        }
      }

      return {
        success: results.length > 0,
        embeddings: results,
        errors: errors,
        model: model
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'GeminiIntegration.generateEmbeddingBatch',
        error: error,
        severity: 'HIGH',
        context: { texts_count: texts ? texts.length : 0, model: model },
        retry: this.isRetryableError(error)
      });

      return {
        success: false,
        embeddings: [],
        errors: [{
          general: error.message
        }]
      };
    }
  }

  /**
   * チャットメッセージからコンテンツを生成します
   * @param {Array<Object>} messages メッセージの配列
   * @param {Object} options 生成オプション
   * @return {Object} 生成結果 {success: boolean, text: string, error: string}
   */
  static async generateChatContent(messages, options = {}) {
    try {
      if (!messages || !Array.isArray(messages) || messages.length === 0) {
        throw new Error('チャットメッセージの配列が無効です');
      }

      // デフォルトオプション
      const config = Config.getSystemConfig();
      const defaultOptions = {
        model: config.generation_model,
        temperature: 0.7,
        maxOutputTokens: 2048,
        systemInstructions: []
      };

      // オプションをマージ
      const mergedOptions = { ...defaultOptions, ...options };

      // API キーを取得
      const apiKey = Config.getAdminGeminiApiKey();
      if (!apiKey) {
        throw new Error('Gemini API キーが構成されていません');
      }

      // メッセージを適切な形式に変換
      const formattedMessages = messages.map(message => {
        return {
          role: message.role || 'user',
          parts: [{ text: message.content || '' }]
        };
      });

      // API リクエストデータを準備
      const requestData = {
        model: mergedOptions.model,
        contents: formattedMessages,
        generationConfig: {
          temperature: mergedOptions.temperature,
          maxOutputTokens: mergedOptions.maxOutputTokens,
          topP: 0.9,
          topK: 40
        }
      };

      // システム指示がある場合は追加
      if (mergedOptions.systemInstructions && mergedOptions.systemInstructions.length > 0) {
        requestData.systemInstructions = {
          parts: mergedOptions.systemInstructions.map(instruction => {
            return { text: instruction };
          })
        };
      }

      // API リクエストを実行
      const response = await this.makeApiRequest('generateContent', requestData, apiKey);

      // レスポンスを検証
      if (!response || !response.candidates || response.candidates.length === 0) {
        throw new Error('API からの応答が無効です');
      }

      const generatedText = response.candidates[0].content.parts[0].text;

      return {
        success: true,
        text: generatedText,
        model: mergedOptions.model
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'GeminiIntegration.generateChatContent',
        error: error,
        severity: 'MEDIUM',
        context: { 
          messages_count: messages ? messages.length : 0,
          model: options.model
        },
        retry: this.isRetryableError(error)
      });

      return {
        success: false,
        error: error.message
      };
    }
  }
}
