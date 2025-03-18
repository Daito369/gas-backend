/**
 * RAG 2.0 応答生成モジュール
 * 検索結果を基にした高品質な応答を生成します
 */
class ResponseGenerator {

  /**
   * 検索結果から応答を生成します
   * @param {Object} searchResults 検索結果オブジェクト（SearchEngine.searchの出力）
   * @param {string} query 元のクエリ
   * @param {Object} options 応答生成オプション
   * @param {string} options.responseType 応答タイプ ('standard', 'email', 'prep', 'detailed')
   * @param {string} options.language 出力言語 ('ja', 'en') - クエリ言語と異なる場合は翻訳
   * @param {string} options.templateId テンプレートID (指定しない場合は自動選択)
   * @param {Object} options.customParams テンプレートに渡すカスタムパラメータ
   * @param {boolean} options.enhanceWithGemini Gemini APIを使用して応答を強化するかどうか
   * @return {Object} 生成された応答 {success: boolean, content: string, error: string}
   */
  static async generateResponse(searchResults, query, options = {}) {
    try {
      // 検索結果の検証
      if (!searchResults || !searchResults.success) {
        throw new Error('検索結果が無効です: ' + (searchResults?.error || '原因不明'));
      }

      // 応答生成のタイミング測定開始
      const startTime = new Date().getTime();

      // デフォルトオプション
      const defaultOptions = {
        responseType: 'standard',
        language: searchResults.language || MultilangProcessor.detectLanguage(query).language || 'ja',
        templateId: null,
        customParams: {},
        enhanceWithGemini: true
      };

      // オプションをマージ
      const mergedOptions = { ...defaultOptions, ...options };

      // チャンク数の確認
      if (!searchResults.results || searchResults.results.length === 0) {
        // 関連ドキュメントが見つからない場合
        return this.generateNoResultsResponse(query, mergedOptions);
      }

      // 応答生成プロセス
      const processedQuery = searchResults.processed_query || query;
      const queryLanguage = searchResults.language || MultilangProcessor.detectLanguage(query).language;
      const topResults = searchResults.results.slice(0, Math.min(5, searchResults.results.length));

      // コンテキスト情報の構築
      const context = await this.buildResponseContext(topResults, processedQuery, queryLanguage);

      // テンプレートの選択または取得
      const template = await this.getResponseTemplate(mergedOptions.templateId, mergedOptions.responseType, context, query);
      
      if (!template) {
        throw new Error('適切なテンプレートが見つかりませんでした');
      }

      // 応答データの構築
      const responseData = {
        query: query,
        processed_query: processedQuery,
        results: topResults,
        context: context,
        params: mergedOptions.customParams,
        language: mergedOptions.language,
        timestamp: new Date().toISOString()
      };

      // テンプレートを適用
      let responseContent = await this.applyTemplate(template, responseData);

      // Gemini APIを使用して応答を強化
      if (mergedOptions.enhanceWithGemini) {
        const enhancedContent = await this.enhanceResponseWithGemini(responseContent, responseData);
        if (enhancedContent.success) {
          responseContent = enhancedContent.content;
        }
      }

      // 出力言語の処理（クエリ言語と異なる場合は翻訳）
      if (mergedOptions.language !== queryLanguage) {
        const translatedContent = await MultilangProcessor.translateText(
          responseContent, queryLanguage, mergedOptions.language
        );
        
        if (translatedContent.success) {
          responseContent = translatedContent.text;
        }
      }

      // 応答生成時間の測定
      const generationTime = new Date().getTime() - startTime;

      // 応答を返す
      return {
        success: true,
        content: responseContent,
        template_id: template.id,
        template_name: template.name,
        response_type: mergedOptions.responseType,
        language: mergedOptions.language,
        generation_time_ms: generationTime
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.generateResponse',
        error: error,
        severity: 'MEDIUM',
        context: { query, options_type: options.responseType },
        retry: false
      });

      // エラー時のフォールバック応答
      return {
        success: false,
        content: this.generateErrorResponse(query, error, options),
        error: error.message
      };
    }
  }

  /**
   * 検索結果がない場合の応答を生成します
   * @param {string} query 検索クエリ
   * @param {Object} options 応答オプション
   * @return {Object} 生成された応答
   */
  static async generateNoResultsResponse(query, options) {
    try {
      // ノーヒット時用のテンプレートを取得
      const templateId = 'no_results_' + options.responseType;
      const template = await this.getResponseTemplate(templateId, 'no_results', null, query);
      
      if (!template) {
        // デフォルトの応答
        const fallbackMessages = {
          'ja': `申し訳ありませんが、「${query}」に関連する情報が見つかりませんでした。以下をお試しください：\n\n・キーワードを変えて検索する\n・より一般的な用語を使用する\n・カテゴリを指定して検索する`,
          'en': `I'm sorry, but I couldn't find any information related to "${query}". Please try:\n\n• Searching with different keywords\n• Using more general terms\n• Specifying a category in your search`
        };
        
        const language = options.language || 'ja';
        return {
          success: true,
          content: fallbackMessages[language] || fallbackMessages['en'],
          response_type: 'no_results',
          language: language
        };
      }
      
      // テンプレートを適用
      const responseData = {
        query: query,
        processed_query: query,
        results: [],
        context: {
          related_queries: this.generateRelatedQueries(query),
          categories: SearchEngine.getCategories()
        },
        params: options.customParams,
        language: options.language,
        timestamp: new Date().toISOString()
      };
      
      const responseContent = await this.applyTemplate(template, responseData);
      
      return {
        success: true,
        content: responseContent,
        template_id: template.id,
        template_name: template.name,
        response_type: 'no_results',
        language: options.language
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.generateNoResultsResponse',
        error: error,
        severity: 'LOW',
        context: { query },
        retry: false
      });
      
      // 最終フォールバック
      return {
        success: true,
        content: `申し訳ありませんが、「${query}」に関連する情報が見つかりませんでした。別のキーワードでお試しください。`,
        response_type: 'no_results',
        language: options.language || 'ja'
      };
    }
  }

  /**
   * エラー発生時の応答を生成します
   * @param {string} query 検索クエリ
   * @param {Error} error エラーオブジェクト
   * @param {Object} options 応答オプション
   * @return {string} エラー応答メッセージ
   */
  static generateErrorResponse(query, error, options = {}) {
    try {
      const language = options.language || 'ja';
      
      // エラータイプ別のユーザーフレンドリーなメッセージを取得
      const userMessage = ErrorHandler.createUserMessage({ 
        message: error.message, 
        severity: 'MEDIUM',
        timestamp: new Date().toISOString()
      }, language);
      
      const errorMessages = {
        'ja': `申し訳ありませんが、「${query}」の処理中にエラーが発生しました。\n\n${userMessage.message}\n\n${userMessage.action}`,
        'en': `I'm sorry, but an error occurred while processing "${query}".\n\n${userMessage.message}\n\n${userMessage.action}`
      };
      
      return errorMessages[language] || errorMessages['en'];
    } catch (e) {
      // エラー処理中のエラー時の最終フォールバック
      return `申し訳ありませんが、応答の生成中にエラーが発生しました。後ほどお試しください。`;
    }
  }

  /**
   * 応答コンテキストを構築します
   * @param {Array<Object>} results 検索結果
   * @param {string} processedQuery 処理済みクエリ
   * @param {string} language 言語
   * @return {Object} 応答コンテキスト
   */
  static async buildResponseContext(results, processedQuery, language) {
    try {
      // ドキュメント情報を収集
      const documentIds = [...new Set(results.map(result => result.document_id))];
      const documents = {};
      
      for (const docId of documentIds) {
        const document = SheetStorage.getDocumentById(docId);
        if (document) {
          documents[docId] = document;
        }
      }
      
      // トピックと概念を抽出
      const topics = this.extractTopicsFromResults(results);
      const concepts = this.extractKeyConceptsFromResults(results);
      
      // 関連ドキュメントをカテゴリ別に整理
      const categorizedDocs = {};
      for (const result of results) {
        const category = result.category || 'general';
        if (!categorizedDocs[category]) {
          categorizedDocs[category] = [];
        }
        if (!categorizedDocs[category].includes(result.document_id)) {
          categorizedDocs[category].push(result.document_id);
        }
      }
      
      // 手順やアクションを抽出
      const procedures = this.extractProceduresFromResults(results);
      const actionItems = this.extractActionItemsFromResults(results);
      
      // 最も関連性の高いスニペットを抽出
      const relevantSnippets = results.slice(0, 3).map(result => ({
        content: result.snippet || result.content.substring(0, 200),
        source: result.title || 'Unknown',
        document_id: result.document_id,
        relevance: result.relevance_score || 0
      }));
      
      // 関連するクエリを生成
      const relatedQueries = this.generateRelatedQueries(processedQuery);
      
      return {
        topics: topics,
        concepts: concepts,
        categorized_documents: categorizedDocs,
        procedures: procedures,
        action_items: actionItems,
        relevant_snippets: relevantSnippets,
        related_queries: relatedQueries,
        language: language,
        documents: documents
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.buildResponseContext',
        error: error,
        severity: 'LOW',
        context: { results_count: results.length },
        retry: false
      });
      
      // エラー時のシンプルなコンテキスト
      return {
        topics: [],
        concepts: [],
        categorized_documents: {},
        procedures: [],
        action_items: [],
        relevant_snippets: results.slice(0, 3).map(result => ({
          content: result.snippet || result.content.substring(0, 200),
          source: result.title || 'Unknown',
          document_id: result.document_id
        })),
        related_queries: [],
        language: language,
        documents: {}
      };
    }
  }

  /**
   * 検索結果からトピックを抽出します
   * @param {Array<Object>} results 検索結果
   * @return {Array<string>} 抽出されたトピック
   */
  static extractTopicsFromResults(results) {
    try {
      const topics = new Set();
      
      // カテゴリを収集
      for (const result of results) {
        if (result.category && result.category !== 'general') {
          topics.add(result.category);
        }
      }
      
      // メタデータからトピックを抽出
      for (const result of results) {
        // ドキュメントメタデータからトピックを抽出
        if (result.document_metadata && result.document_metadata.topics) {
          for (const topic of result.document_metadata.topics) {
            topics.add(topic);
          }
        }
        
        // チャンクメタデータからトピックを抽出
        if (result.metadata && result.metadata.topics) {
          for (const topic of result.metadata.topics) {
            topics.add(topic);
          }
        }
      }
      
      // 頻出語をトピックとして抽出
      const contentText = results.map(r => r.content || r.snippet || '').join(' ');
      const words = contentText.toLowerCase()
        .replace(/[^\w\s]/g, ' ')
        .split(/\s+/)
        .filter(word => word.length > 3);
      
      const wordCounts = {};
      for (const word of words) {
        wordCounts[word] = (wordCounts[word] || 0) + 1;
      }
      
      // 頻出上位5単語を抽出
      const topWords = Object.entries(wordCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .map(entry => entry[0]);
      
      for (const word of topWords) {
        topics.add(word);
      }
      
      return Array.from(topics);
    } catch (error) {
      return [];
    }
  }

  /**
   * 検索結果から主要概念を抽出します
   * @param {Array<Object>} results 検索結果
   * @return {Array<Object>} 抽出された概念
   */
  static extractKeyConceptsFromResults(results) {
    try {
      const concepts = {};
      
      // メタデータから概念を抽出
      for (const result of results) {
        // ドキュメントメタデータから概念を抽出
        if (result.document_metadata && result.document_metadata.concepts) {
          for (const concept of result.document_metadata.concepts) {
            concepts[concept.name] = {
              name: concept.name,
              description: concept.description || '',
              relevance: concept.relevance || 1
            };
          }
        }
        
        // チャンクメタデータから概念を抽出
        if (result.metadata && result.metadata.concepts) {
          for (const concept of result.metadata.concepts) {
            concepts[concept.name] = {
              name: concept.name,
              description: concept.description || '',
              relevance: concept.relevance || 1
            };
          }
        }
      }
      
      // マッチしたキーワードを概念として追加
      for (const result of results) {
        if (result.matched_keywords && result.matched_keywords.length > 0) {
          for (const keyword of result.matched_keywords) {
            if (!concepts[keyword] && keyword.length > 3) {
              // コンテンツから該当するキーワードの周辺テキストを抽出
              const description = this.extractContextForKeyword(result.content, keyword);
              
              concepts[keyword] = {
                name: keyword,
                description: description,
                relevance: 0.7
              };
            }
          }
        }
      }
      
      // 概念を配列に変換し、関連性でソート
      return Object.values(concepts).sort((a, b) => b.relevance - a.relevance);
    } catch (error) {
      return [];
    }
  }

  /**
   * キーワードの周辺コンテキストを抽出します
   * @param {string} content コンテンツ
   * @param {string} keyword キーワード
   * @return {string} 周辺コンテキスト
   */
  static extractContextForKeyword(content, keyword) {
    if (!content || !keyword) return '';
    
    try {
      const contentLower = content.toLowerCase();
      const keywordLower = keyword.toLowerCase();
      const index = contentLower.indexOf(keywordLower);
      
      if (index === -1) return '';
      
      // キーワードの前後のコンテキストを抽出
      const startIndex = Math.max(0, index - 50);
      const endIndex = Math.min(content.length, index + keyword.length + 50);
      let context = content.substring(startIndex, endIndex);
      
      // 文が途中で切れている場合は調整
      if (startIndex > 0) {
        const firstSpaceIndex = context.indexOf(' ');
        if (firstSpaceIndex > 0) {
          context = context.substring(firstSpaceIndex + 1);
        }
      }
      
      if (endIndex < content.length) {
        const lastSpaceIndex = context.lastIndexOf(' ');
        if (lastSpaceIndex > 0) {
          context = context.substring(0, lastSpaceIndex);
        }
      }
      
      return context;
    } catch (error) {
      return '';
    }
  }

  /**
   * 検索結果から手順を抽出します
   * @param {Array<Object>} results 検索結果
   * @return {Array<Object>} 抽出された手順
   */
  static extractProceduresFromResults(results) {
    try {
      const procedures = [];
      
      for (const result of results) {
        // 番号付きリストや箇条書きパターンを検出
        const content = result.content || '';
        
        // 番号付きリストの検出
        const numberedListPattern = /^\s*(\d+\.\s+.+)(?:\n\s*\d+\.\s+.+)*$/gm;
        const numberedLists = content.match(numberedListPattern);
        
        if (numberedLists) {
          for (const list of numberedLists) {
            const steps = list.split('\n')
              .map(step => step.trim())
              .filter(step => /^\d+\.\s+.+$/.test(step))
              .map(step => step.replace(/^\d+\.\s+/, ''));
            
            if (steps.length >= 2) {
              procedures.push({
                title: `From ${result.title || 'document'}`,
                steps: steps,
                source: result.document_id,
                type: 'numbered_list'
              });
            }
          }
        }
        
        // 箇条書きの検出
        const bulletListPattern = /^\s*([•\-*]\s+.+)(?:\n\s*[•\-*]\s+.+)*$/gm;
        const bulletLists = content.match(bulletListPattern);
        
        if (bulletLists) {
          for (const list of bulletLists) {
            const steps = list.split('\n')
              .map(step => step.trim())
              .filter(step => /^[•\-*]\s+.+$/.test(step))
              .map(step => step.replace(/^[•\-*]\s+/, ''));
            
            if (steps.length >= 2) {
              procedures.push({
                title: `From ${result.title || 'document'}`,
                steps: steps,
                source: result.document_id,
                type: 'bullet_list'
              });
            }
          }
        }
        
        // メタデータから手順を抽出
        if (result.metadata && result.metadata.procedures) {
          for (const procedure of result.metadata.procedures) {
            procedures.push({
              title: procedure.title || `From ${result.title || 'document'}`,
              steps: procedure.steps || [],
              source: result.document_id,
              type: 'metadata'
            });
          }
        }
      }
      
      return procedures;
    } catch (error) {
      return [];
    }
  }

  /**
   * 検索結果からアクションアイテムを抽出します
   * @param {Array<Object>} results 検索結果
   * @return {Array<string>} 抽出されたアクションアイテム
   */
  static extractActionItemsFromResults(results) {
    try {
      const actionItems = new Set();
      
      // アクション指向の文を探す
      const actionPhrases = [
        /you should/i, /you need to/i, /please/i, /required/i, /must/i, /important to/i,
        /必要があります/i, /してください/i, /必須です/i, /重要です/i, /行ってください/i
      ];
      
      for (const result of results) {
        const content = result.content || '';
        const sentences = content.split(/[.。]/).map(s => s.trim()).filter(s => s.length > 10);
        
        for (const sentence of sentences) {
          // アクション指向の文を検出
          if (actionPhrases.some(phrase => phrase.test(sentence))) {
            actionItems.add(sentence);
          }
        }
        
        // メタデータからアクションアイテムを抽出
        if (result.metadata && result.metadata.action_items) {
          for (const action of result.metadata.action_items) {
            actionItems.add(action);
          }
        }
      }
      
      return Array.from(actionItems);
    } catch (error) {
      return [];
    }
  }

  /**
   * 関連クエリを生成します
   * @param {string} query 元のクエリ
   * @return {Array<string>} 関連クエリの配列
   */
  static generateRelatedQueries(query) {
    try {
      // 簡易的な関連クエリ生成
      const relatedQueries = [];
      
      // 1. プレフィックスを追加したクエリ
      const prefixes = ['how to', 'what is', 'troubleshoot', 'guide for'];
      const words = query.split(/\s+/);
      
      if (words.length <= 3) {
        for (const prefix of prefixes) {
          if (!query.toLowerCase().startsWith(prefix)) {
            relatedQueries.push(`${prefix} ${query}`);
          }
        }
      }
      
      // 2. サフィックスを追加したクエリ
      const suffixes = ['examples', 'tutorial', 'guide', 'best practices'];
      for (const suffix of suffixes) {
        if (!query.toLowerCase().endsWith(suffix)) {
          relatedQueries.push(`${query} ${suffix}`);
        }
      }
      
      // 3. 特定の単語を置換したクエリ
      const keywords = words.filter(word => word.length > 3);
      const synonyms = {
        'error': ['issue', 'problem', 'trouble'],
        'setup': ['configuration', 'settings', 'set up'],
        'guide': ['tutorial', 'instructions', 'how-to'],
        'change': ['modify', 'update', 'edit'],
        'report': ['analytics', 'statistics', 'data']
      };
      
      for (const keyword of keywords) {
        const lowerKeyword = keyword.toLowerCase();
        for (const [word, alternatives] of Object.entries(synonyms)) {
          if (lowerKeyword === word) {
            for (const alternative of alternatives) {
              const newQuery = query.replace(new RegExp(`\\b${keyword}\\b`, 'i'), alternative);
              relatedQueries.push(newQuery);
            }
          }
        }
      }
      
      // 配列を返す前に重複を削除してシャッフル
      return this.shuffleArray([...new Set(relatedQueries)]).slice(0, 3);
    } catch (error) {
      return [];
    }
  }

  /**
   * 配列をシャッフルします
   * @param {Array} array シャッフルする配列
   * @return {Array} シャッフルされた配列
   */
  static shuffleArray(array) {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
  }

  /**
   * 応答テンプレートを取得します
   * @param {string} templateId テンプレートID
   * @param {string} responseType 応答タイプ
   * @param {Object} context コンテキスト情報
   * @param {string} query 検索クエリ
   * @return {Object} テンプレートオブジェクト
   */
  static async getResponseTemplate(templateId, responseType, context, query) {
    try {
      // キャッシュをチェック
      const cacheKey = `template_${templateId || responseType}_${Utilities.generateUniqueId().substring(0, 8)}`;
      const cachedTemplate = CacheManager.get(cacheKey);
      
      if (cachedTemplate) {
        return JSON.parse(cachedTemplate);
      }
      
      // IDでテンプレートを取得
      if (templateId) {
        return await this.getTemplateById(templateId);
      }
      
      // レスポンスタイプとコンテキストに基づいてテンプレートを選択
      const template = await this.selectTemplate(responseType, context, query);
      
      // キャッシュに保存（1時間）
      if (template) {
        CacheManager.set(cacheKey, JSON.stringify(template), 3600);
      }
      
      return template;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.getResponseTemplate',
        error: error,
        severity: 'MEDIUM',
        context: { template_id: templateId, response_type: responseType },
        retry: false
      });
      
      // エラー時のデフォルトテンプレート
      return this.getDefaultTemplate(responseType);
    }
  }

  /**
   * IDからテンプレートを取得します
   * @param {string} templateId テンプレートID
   * @return {Object} テンプレートオブジェクト
   */
  static async getTemplateById(templateId) {
    try {
      // キャッシュをチェック
      const cacheKey = `template_by_id_${templateId}`;
      const cachedTemplate = CacheManager.get(cacheKey);
      
      if (cachedTemplate) {
        return JSON.parse(cachedTemplate);
      }
      
      // スプレッドシートからテンプレートを取得
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const templatesSheet = ss.getSheetByName('Templates');
      
      if (!templatesSheet) {
        throw new Error('Templates シートが見つかりません');
      }
      
      const data = templatesSheet.getDataRange().getValues();
      if (data.length <= 1) {
        throw new Error('テンプレートデータが見つかりません');
      }
      
      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const idIndex = headers.indexOf('id');
      const nameIndex = headers.indexOf('name');
      const typeIndex = headers.indexOf('type');
      const contentIndex = headers.indexOf('content');
      const languageIndex = headers.indexOf('language');
      const categoryIndex = headers.indexOf('category');
      const metadataIndex = headers.indexOf('metadata');
      
      // 必須フィールドの検証
      if (idIndex === -1 || contentIndex === -1) {
        throw new Error('Templates シートの形式が不正です');
      }
      
      // テンプレートIDに一致する行を検索
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[idIndex] === templateId) {
          const template = {
            id: row[idIndex],
            name: nameIndex !== -1 ? row[nameIndex] : '',
            type: typeIndex !== -1 ? row[typeIndex] : 'standard',
            content: row[contentIndex],
            language: languageIndex !== -1 ? row[languageIndex] : 'ja',
            category: categoryIndex !== -1 ? row[categoryIndex] : '',
            metadata: metadataIndex !== -1 ? this.parseMetadata(row[metadataIndex]) : {}
          };
          
          // キャッシュに保存（1日）
          CacheManager.set(cacheKey, JSON.stringify(template), 86400);
          
          return template;
        }
      }
      
      throw new Error(`テンプレート ${templateId} が見つかりません`);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.getTemplateById',
        error: error,
        severity: 'MEDIUM',
        context: { template_id: templateId },
        retry: false
      });
      
      return null;
    }
  }

  /**
   * コンテキストと応答タイプに最適なテンプレートを選択します
   * @param {string} responseType 応答タイプ
   * @param {Object} context コンテキスト情報
   * @param {string} query 検索クエリ
   * @return {Object} 選択されたテンプレート
   */
  static async selectTemplate(responseType, context, query) {
    try {
      // すべてのテンプレートをロード
      const templates = await this.loadTemplates();
      
      if (!templates || templates.length === 0) {
        return this.getDefaultTemplate(responseType);
      }
      
      // 応答タイプに一致するテンプレートを絞り込み
      let candidates = templates.filter(template => template.type === responseType);
      
      if (candidates.length === 0) {
        candidates = templates.filter(template => template.type === 'standard');
      }
      
      if (candidates.length === 0) {
        return this.getDefaultTemplate(responseType);
      }
      
      // コンテキストがない場合は最初のテンプレートを返す
      if (!context) {
        return candidates[0];
      }
      
      // 言語に一致するテンプレートを優先
      const languageMatches = candidates.filter(template => template.language === context.language);
      
      if (languageMatches.length > 0) {
        candidates = languageMatches;
      }
      
      // カテゴリに一致するテンプレートを優先
      if (context.categorized_documents) {
        const categories = Object.keys(context.categorized_documents);
        
        if (categories.length > 0) {
          const categoryMatches = candidates.filter(template => 
            template.category && categories.includes(template.category)
          );
          
          if (categoryMatches.length > 0) {
            candidates = categoryMatches;
          }
        }
      }
      
      // 応答タイプ別の特殊な選択ロジック
      switch (responseType) {
        case 'email':
          // メールテンプレートの選択ロジック
          return this.selectEmailTemplate(candidates, context, query);
          
        case 'prep':
          // PREPテンプレートの選択ロジック
          return this.selectPrepTemplate(candidates, context, query);
          
        case 'detailed':
          // 詳細テンプレートの選択ロジック
          return this.selectDetailedTemplate(candidates, context, query);
          
        default:
          // 標準テンプレートの選択ロジック
          // 最初のテンプレートを返す
          return candidates[0];
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.selectTemplate',
        error: error,
        severity: 'MEDIUM',
        context: { response_type: responseType },
        retry: false
      });
      
      return this.getDefaultTemplate(responseType);
    }
  }

  /**
   * メールテンプレートを選択します
   * @param {Array<Object>} candidates 候補テンプレート
   * @param {Object} context コンテキスト情報
   * @param {string} query 検索クエリ
   * @return {Object} 選択されたテンプレート
   */
  static selectEmailTemplate(candidates, context, query) {
    // クエリから緊急度と複雑さを評価
    const urgencyPatterns = [
      /urgent/i, /asap/i, /immediately/i, /emergency/i,
      /緊急/i, /早急/i, /すぐに/i, /今すぐ/i
    ];
    
    const urgency = urgencyPatterns.some(pattern => pattern.test(query)) ? 'high' : 'normal';
    
    // 複雑さの評価（手順の数とスニペットの長さに基づく）
    let complexity = 'simple';
    if (context.procedures && context.procedures.length > 0) {
      const totalSteps = context.procedures.reduce((sum, proc) => sum + proc.steps.length, 0);
      if (totalSteps > 10) {
        complexity = 'complex';
      } else if (totalSteps > 5) {
        complexity = 'moderate';
      }
    }
    
    // テンプレートのメタデータを確認
    for (const template of candidates) {
      const metadata = template.metadata || {};
      
      // 緊急度と複雑さに一致するテンプレートを探す
      if (metadata.urgency === urgency && metadata.complexity === complexity) {
        return template;
      }
    }
    
    // 緊急度のみ一致するテンプレートを探す
    for (const template of candidates) {
      const metadata = template.metadata || {};
      if (metadata.urgency === urgency) {
        return template;
      }
    }
    
    // 複雑さのみ一致するテンプレートを探す
    for (const template of candidates) {
      const metadata = template.metadata || {};
      if (metadata.complexity === complexity) {
        return template;
      }
    }
    
    // 一致するものがなければ最初のテンプレートを返す
    return candidates[0];
  }

  /**
   * PREPテンプレートを選択します
   * @param {Array<Object>} candidates 候補テンプレート
   * @param {Object} context コンテキスト情報
   * @param {string} query 検索クエリ
   * @return {Object} 選択されたテンプレート
   */
  static selectPrepTemplate(candidates, context, query) {
    // クエリからタイプを評価
    const queryLower = query.toLowerCase();
    
    let prepType = 'general';
    
    // トラブルシューティング
    if (/error|issue|problem|troubleshoot|fix|resolve|エラー|問題|トラブル|解決/.test(queryLower)) {
      prepType = 'troubleshooting';
    }
    // 機能説明
    else if (/how|what|explain|guide|setup|configure|方法|説明|ガイド|設定/.test(queryLower)) {
      prepType = 'explanation';
    }
    // ポリシー説明
    else if (/policy|rule|term|condition|compliant|ポリシー|規約|条件|コンプライアンス/.test(queryLower)) {
      prepType = 'policy';
    }
    
    // テンプレートのメタデータを確認
    for (const template of candidates) {
      const metadata = template.metadata || {};
      
      // タイプに一致するテンプレートを探す
      if (metadata.prep_type === prepType) {
        return template;
      }
    }
    
    // 一致するものがなければ最初のテンプレートを返す
    return candidates[0];
  }

  /**
   * 詳細テンプレートを選択します
   * @param {Array<Object>} candidates 候補テンプレート
   * @param {Object} context コンテキスト情報
   * @param {string} query 検索クエリ
   * @return {Object} 選択されたテンプレート
   */
  static selectDetailedTemplate(candidates, context, query) {
    // カテゴリベースで選択
    if (context.categorized_documents) {
      const categories = Object.keys(context.categorized_documents);
      
      if (categories.length > 0) {
        // 最も関連性の高いカテゴリ（最初のカテゴリ）
        const primaryCategory = categories[0];
        
        for (const template of candidates) {
          if (template.category === primaryCategory) {
            return template;
          }
        }
      }
    }
    
    // 詳細度に基づく選択
    let detailLevel = 'standard';
    
    // クエリから詳細度を評価
    if (/detailed|complete|full|comprehensive|詳細|完全|全面的/.test(query.toLowerCase())) {
      detailLevel = 'high';
    }
    
    // テンプレートのメタデータを確認
    for (const template of candidates) {
      const metadata = template.metadata || {};
      
      // 詳細度に一致するテンプレートを探す
      if (metadata.detail_level === detailLevel) {
        return template;
      }
    }
    
    // 一致するものがなければ最初のテンプレートを返す
    return candidates[0];
  }

  /**
   * すべてのテンプレートをロードします
   * @return {Array<Object>} テンプレートの配列
   */
  static async loadTemplates() {
    try {
      // キャッシュをチェック
      const cacheKey = 'all_templates';
      const cachedTemplates = CacheManager.get(cacheKey);
      
      if (cachedTemplates) {
        return JSON.parse(cachedTemplates);
      }
      
      // スプレッドシートからテンプレートをロード
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const templatesSheet = ss.getSheetByName('Templates');
      
      if (!templatesSheet) {
        throw new Error('Templates シートが見つかりません');
      }
      
      const data = templatesSheet.getDataRange().getValues();
      if (data.length <= 1) {
        throw new Error('テンプレートデータが見つかりません');
      }
      
      // ヘッダー行を取得してインデックスをマッピング
      const headers = data[0];
      const idIndex = headers.indexOf('id');
      const nameIndex = headers.indexOf('name');
      const typeIndex = headers.indexOf('type');
      const contentIndex = headers.indexOf('content');
      const languageIndex = headers.indexOf('language');
      const categoryIndex = headers.indexOf('category');
      const metadataIndex = headers.indexOf('metadata');
      
      // 必須フィールドの検証
      if (idIndex === -1 || contentIndex === -1) {
        throw new Error('Templates シートの形式が不正です');
      }
      
      // テンプレートを収集
      const templates = [];
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const template = {
          id: row[idIndex],
          name: nameIndex !== -1 ? row[nameIndex] : '',
          type: typeIndex !== -1 ? row[typeIndex] : 'standard',
          content: row[contentIndex],
          language: languageIndex !== -1 ? row[languageIndex] : 'ja',
          category: categoryIndex !== -1 ? row[categoryIndex] : '',
          metadata: metadataIndex !== -1 ? this.parseMetadata(row[metadataIndex]) : {}
        };
        
        templates.push(template);
      }
      
      // キャッシュに保存（30分）
      CacheManager.set(cacheKey, JSON.stringify(templates), 1800);
      
      return templates;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.loadTemplates',
        error: error,
        severity: 'MEDIUM',
        retry: false
      });
      
      return [];
    }
  }

  /**
   * テンプレートを応答データに適用します
   * @param {Object} template テンプレートオブジェクト
   * @param {Object} responseData 応答データ
   * @return {string} フォーマットされた応答
   */
  static async applyTemplate(template, responseData) {
    try {
      let content = template.content;
      
      // 基本的な置換を実行
      content = this.replaceBasicPlaceholders(content, responseData);
      
      // 条件付きブロックを処理
      content = this.processConditionalBlocks(content, responseData);
      
      // 繰り返しブロックを処理
      content = this.processIterationBlocks(content, responseData);
      
      // フォーマット関数を処理
      content = this.processFormatFunctions(content, responseData);
      
      return content;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.applyTemplate',
        error: error,
        severity: 'MEDIUM',
        context: { template_id: template.id },
        retry: false
      });
      
      // エラー時のシンプルな応答
      return this.createSimpleResponse(responseData);
    }
  }

  /**
   * 基本的なプレースホルダを置換します
   * @param {string} content テンプレートコンテンツ
   * @param {Object} data 応答データ
   * @return {string} 置換後のコンテンツ
   */
  static replaceBasicPlaceholders(content, data) {
    // クエリの置換
    content = content.replace(/\{query\}/g, data.query || '');
    content = content.replace(/\{processed_query\}/g, data.processed_query || data.query || '');
    
    // 言語の置換
    content = content.replace(/\{language\}/g, data.language || 'ja');
    
    // タイムスタンプの置換
    content = content.replace(/\{timestamp\}/g, data.timestamp || new Date().toISOString());
    content = content.replace(/\{date\}/g, new Date().toLocaleDateString());
    content = content.replace(/\{time\}/g, new Date().toLocaleTimeString());
    
    // カスタムパラメータの置換
    if (data.params) {
      for (const [key, value] of Object.entries(data.params)) {
        content = content.replace(new RegExp(`\\{param\\.${key}\\}`, 'g'), value);
      }
    }
    
    return content;
  }

  /**
   * 条件付きブロックを処理します
   * @param {string} content テンプレートコンテンツ
   * @param {Object} data 応答データ
   * @return {string} 処理後のコンテンツ
   */
  static processConditionalBlocks(content, data) {
    // {if condition}...{else}...{endif} ブロックの処理
    const ifRegex = /\{if\s+([^}]+)\}([\s\S]*?)(?:\{else\}([\s\S]*?))?\{endif\}/g;
    
    return content.replace(ifRegex, (match, condition, ifContent, elseContent = '') => {
      let result = false;
      
      try {
        // 条件の評価
        if (condition.includes('==')) {
          const [left, right] = condition.split('==').map(s => s.trim());
          const leftValue = this.getValueFromPath(data, left);
          const rightValue = right.startsWith('"') && right.endsWith('"') 
            ? right.slice(1, -1) 
            : this.getValueFromPath(data, right);
          
          result = leftValue == rightValue; // 弱い等価性を使用
        }
        else if (condition.includes('!=')) {
          const [left, right] = condition.split('!=').map(s => s.trim());
          const leftValue = this.getValueFromPath(data, left);
          const rightValue = right.startsWith('"') && right.endsWith('"') 
            ? right.slice(1, -1) 
            : this.getValueFromPath(data, right);
          
          result = leftValue != rightValue; // 弱い不等価性を使用
        }
        else if (condition.includes('>')) {
          const [left, right] = condition.split('>').map(s => s.trim());
          const leftValue = Number(this.getValueFromPath(data, left));
          const rightValue = Number(right.startsWith('"') && right.endsWith('"') 
            ? right.slice(1, -1) 
            : this.getValueFromPath(data, right));
          
          result = !isNaN(leftValue) && !isNaN(rightValue) && leftValue > rightValue;
        }
        else if (condition.includes('<')) {
          const [left, right] = condition.split('<').map(s => s.trim());
          const leftValue = Number(this.getValueFromPath(data, left));
          const rightValue = Number(right.startsWith('"') && right.endsWith('"') 
            ? right.slice(1, -1) 
            : this.getValueFromPath(data, right));
          
          result = !isNaN(leftValue) && !isNaN(rightValue) && leftValue < rightValue;
        }
        else if (condition.includes('exists')) {
          const path = condition.replace('exists', '').trim();
          const value = this.getValueFromPath(data, path);
          result = value !== undefined && value !== null && 
                 (typeof value !== 'object' || Object.keys(value).length > 0) &&
                 (Array.isArray(value) ? value.length > 0 : true);
        }
        else if (condition.includes('empty')) {
          const path = condition.replace('empty', '').trim();
          const value = this.getValueFromPath(data, path);
          result = value === undefined || value === null || value === '' ||
                 (typeof value === 'object' && Object.keys(value).length === 0) ||
                 (Array.isArray(value) && value.length === 0);
        }
        else {
          // シンプルな存在チェック
          const value = this.getValueFromPath(data, condition);
          result = value !== undefined && value !== null && value !== '';
        }
      } catch (e) {
        // 条件評価エラー時はelseコンテンツを使用
        result = false;
      }
      
      return result ? ifContent : elseContent;
    });
  }

  /**
   * 繰り返しブロックを処理します
   * @param {string} content テンプレートコンテンツ
   * @param {Object} data 応答データ
   * @return {string} 処理後のコンテンツ
   */
  static processIterationBlocks(content, data) {
    // {for item in collection}...{endfor} ブロックの処理
    const forRegex = /\{for\s+(\w+)\s+in\s+([^}]+)\}([\s\S]*?)\{endfor\}/g;
    
    return content.replace(forRegex, (match, itemName, collectionPath, itemTemplate) => {
      try {
        const collection = this.getValueFromPath(data, collectionPath);
        
        if (!collection || !Array.isArray(collection) || collection.length === 0) {
          return ''; // コレクションがない場合は空文字列を返す
        }
        
        let result = '';
        
        for (let i = 0; i < collection.length; i++) {
          const item = collection[i];
          let itemContent = itemTemplate;
          
          // {item} プレースホルダーの置換
          itemContent = itemContent.replace(new RegExp(`\\{${itemName}\\}`, 'g'), 
            typeof item === 'object' ? JSON.stringify(item) : item);
          
          // {item.property} プレースホルダーの置換
          if (typeof item === 'object') {
            for (const [key, value] of Object.entries(item)) {
              itemContent = itemContent.replace(
                new RegExp(`\\{${itemName}\\.${key}\\}`, 'g'), 
                typeof value === 'object' ? JSON.stringify(value) : value
              );
            }
          }
          
          // {index} プレースホルダーの置換
          itemContent = itemContent.replace(/\{index\}/g, i + 1);
          
          result += itemContent;
        }
        
        return result;
      } catch (e) {
        return ''; // エラー時は空文字列を返す
      }
    });
  }

  /**
   * フォーマット関数を処理します
   * @param {string} content テンプレートコンテンツ
   * @param {Object} data 応答データ
   * @return {string} 処理後のコンテンツ
   */
  static processFormatFunctions(content, data) {
    // {format:function(path)} の処理
    const formatRegex = /\{format:(\w+)\(([^)]+)\)\}/g;
    
    return content.replace(formatRegex, (match, functionName, path) => {
      try {
        const value = this.getValueFromPath(data, path.trim());
        
        switch (functionName) {
          case 'date':
            if (value) {
              const date = new Date(value);
              return date.toLocaleDateString();
            }
            return '';
            
          case 'time':
            if (value) {
              const date = new Date(value);
              return date.toLocaleTimeString();
            }
            return '';
            
          case 'datetime':
            if (value) {
              const date = new Date(value);
              return date.toLocaleString();
            }
            return '';
            
          case 'upper':
            return value ? String(value).toUpperCase() : '';
            
          case 'lower':
            return value ? String(value).toLowerCase() : '';
            
          case 'capitalize':
            return value ? String(value).charAt(0).toUpperCase() + String(value).slice(1) : '';
            
          case 'number':
            return value ? Number(value).toLocaleString() : '';
            
          case 'json':
            return value ? JSON.stringify(value, null, 2) : '';
            
          case 'list':
            if (Array.isArray(value) && value.length > 0) {
              return value.map(item => `- ${item}`).join('\n');
            }
            return '';
            
          case 'count':
            if (Array.isArray(value)) {
              return value.length.toString();
            }
            return '0';
            
          case 'truncate':
            const maxLength = path.split(',')[1] ? parseInt(path.split(',')[1].trim()) : 100;
            if (value && typeof value === 'string' && value.length > maxLength) {
              return value.substring(0, maxLength) + '...';
            }
            return value || '';
            
          default:
            return value || '';
        }
      } catch (e) {
        return ''; // エラー時は空文字列を返す
      }
    });
  }

  /**
   * データパスから値を取得します
   * @param {Object} data データオブジェクト
   * @param {string} path ドット区切りのパス
   * @return {*} パスで指定された値
   */
  static getValueFromPath(data, path) {
    if (!data || !path) return undefined;
    
    const parts = path.split('.');
    let value = data;
    
    for (const part of parts) {
      if (value === null || value === undefined) return undefined;
      
      // 配列インデックスの処理
      if (part.includes('[') && part.includes(']')) {
        const name = part.substring(0, part.indexOf('['));
        const index = parseInt(part.substring(part.indexOf('[') + 1, part.indexOf(']')));
        
        if (isNaN(index)) return undefined;
        
        value = value[name];
        if (!Array.isArray(value)) return undefined;
        
        value = value[index];
      } else {
        value = value[part];
      }
    }
    
    return value;
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
   * デフォルトテンプレートを取得します
   * @param {string} responseType 応答タイプ
   * @return {Object} テンプレートオブジェクト
   */
  static getDefaultTemplate(responseType) {
    let templateContent = '';
    
    // 応答タイプに基づいてデフォルトテンプレートを生成
    switch (responseType) {
      case 'email':
        templateContent = this.getDefaultEmailTemplate();
        break;
        
      case 'prep':
        templateContent = this.getDefaultPrepTemplate();
        break;
        
      case 'detailed':
        templateContent = this.getDefaultDetailedTemplate();
        break;
        
      case 'no_results':
        templateContent = this.getDefaultNoResultsTemplate();
        break;
        
      case 'standard':
      default:
        templateContent = this.getDefaultStandardTemplate();
        break;
    }
    
    // テンプレートオブジェクトを作成
    return {
      id: `default_${responseType}`,
      name: `Default ${responseType} Template`,
      type: responseType,
      content: templateContent,
      language: 'ja', // デフォルト言語
      category: '',
      metadata: {}
    };
  }

  /**
   * デフォルトの標準テンプレートを取得します
   * @return {string} テンプレートコンテンツ
   */
  static getDefaultStandardTemplate() {
    return `### 「{query}」に関する回答

{if exists context.relevant_snippets}
検索結果から以下の情報が見つかりました：

{for snippet in context.relevant_snippets}
> {snippet.content}
*出典: {snippet.source}*

{endfor}
{endif}

{if exists context.procedures}
**手順:**

{for procedure in context.procedures}
{procedure.title}:
{for step in procedure.steps}
{index}. {step}
{endfor}

{endfor}
{endif}

{if exists context.action_items}
**推奨アクション:**

{for action in context.action_items}
- {action}
{endfor}
{endif}

これは自動生成された回答です。ご不明な点がございましたら、お気軽にお問い合わせください。`;
  }

  /**
   * デフォルトのメールテンプレートを取得します
   * @return {string} テンプレートコンテンツ
   */
  static getDefaultEmailTemplate() {
    return `件名: Re: {query}

お世話になっております。
Google広告サポートチームです。

ご質問いただきました「{query}」について回答いたします。

{if exists context.relevant_snippets}
{for snippet in context.relevant_snippets}
{snippet.content}

{endfor}
{endif}

{if exists context.procedures}
■ 手順
{for procedure in context.procedures}
【{procedure.title}】
{for step in procedure.steps}
{index}. {step}
{endfor}

{endfor}
{endif}

{if exists context.action_items}
■ 推奨アクション
{for action in context.action_items}
・{action}
{endfor}
{endif}

ご不明な点がございましたら、お気軽にご返信ください。
引き続きよろしくお願いいたします。

--
Google広告サポートチーム`;
  }

  /**
   * デフォルトのPREPテンプレートを取得します
   * @return {string} テンプレートコンテンツ
   */
  static getDefaultPrepTemplate() {
    return `# PREP: {query}

## Point（要点）
{if exists context.relevant_snippets}
{context.relevant_snippets[0].content}
{endif}

## Reason（理由）
{if exists context.concepts}
{for concept in context.concepts}
- {concept.name}: {concept.description}
{endfor}
{endif}

## Example（例）
{if exists context.procedures}
{for procedure in context.procedures}
### {procedure.title}
{for step in procedure.steps}
{index}. {step}
{endfor}
{endfor}
{endif}

## Proposal（提案）
{if exists context.action_items}
{for action in context.action_items}
- {action}
{endfor}
{endif}`;
  }

  /**
   * デフォルトの詳細テンプレートを取得します
   * @return {string} テンプレートコンテンツ
   */
  static getDefaultDetailedTemplate() {
    return `# 「{query}」に関する詳細情報

## 概要
{if exists context.relevant_snippets}
{context.relevant_snippets[0].content}
{endif}

## 詳細説明
{if exists context.relevant_snippets}
{for snippet in context.relevant_snippets}
{snippet.content}

*出典: {snippet.source}*

{endfor}
{endif}

## 主要概念
{if exists context.concepts}
{for concept in context.concepts}
### {concept.name}
{concept.description}

{endfor}
{endif}

## 手順とガイド
{if exists context.procedures}
{for procedure in context.procedures}
### {procedure.title}
{for step in procedure.steps}
{index}. {step}
{endfor}

{endfor}
{endif}

## 推奨アクション
{if exists context.action_items}
{for action in context.action_items}
- {action}
{endfor}
{endif}

## 関連トピック
{if exists context.topics}
{for topic in context.topics}
- {topic}
{endfor}
{endif}

{if exists context.related_queries}
## 関連検索
{for query in context.related_queries}
- {query}
{endfor}
{endif}`;
  }

  /**
   * デフォルトの検索結果なしテンプレートを取得します
   * @return {string} テンプレートコンテンツ
   */
  static getDefaultNoResultsTemplate() {
    return `### 「{query}」に関する情報が見つかりませんでした

申し訳ありませんが、お探しの情報が見つかりませんでした。以下をお試しください：

- キーワードを変えて検索する
- より一般的な用語を使用する
- カテゴリを指定して検索する

{if exists context.related_queries}
**別の検索候補:**
{for query in context.related_queries}
- {query}
{endfor}
{endif}

{if exists context.categories}
**カテゴリ検索:**
{for category in context.categories}
- {category}
{endfor}
{endif}`;
  }

  /**
   * シンプルな応答を作成します
   * @param {Object} responseData 応答データ
   * @return {string} シンプルな応答
   */
  static createSimpleResponse(responseData) {
    // 検索結果がある場合
    if (responseData.results && responseData.results.length > 0) {
      let response = `### 「${responseData.query}」に関する回答\n\n`;
      
      // 上位3件の結果を表示
      const topResults = responseData.results.slice(0, 3);
      response += '検索結果から以下の情報が見つかりました：\n\n';
      
      for (const result of topResults) {
        response += `> ${result.snippet || result.content.substring(0, 200)}\n`;
        response += `*出典: ${result.title || 'ドキュメント'}*\n\n`;
      }
      
      response += 'これは自動生成された回答です。ご不明な点がございましたら、お気軽にお問い合わせください。';
      
      return response;
    }
    
    // 検索結果がない場合
    return `### 「${responseData.query}」に関する情報が見つかりませんでした\n\n申し訳ありませんが、お探しの情報が見つかりませんでした。別のキーワードで検索してみてください。`;
  }

  /**
   * Gemini APIを使用して応答を強化します
   * @param {string} responseContent 応答コンテンツ
   * @param {Object} responseData 応答データ
   * @return {Object} 強化された応答 {success: boolean, content: string}
   */
  static async enhanceResponseWithGemini(responseContent, responseData) {
    try {
      // レスポンスがすでに十分に長い場合は強化しない
      if (responseContent.length > 1000) {
        return {
          success: true,
          content: responseContent
        };
      }
      
      // 検索結果がない場合は強化しない
      if (!responseData.results || responseData.results.length === 0) {
        return {
          success: true,
          content: responseContent
        };
      }
      
      // Gemini APIへのプロンプト作成
      const language = responseData.language || 'ja';
      
      const promptParts = [
        language === 'ja' 
          ? '以下の情報を基に、より完成された回答を生成してください。元の回答構造を維持し、検索結果から得られた情報を適切に組み込んでください。情報がない場合は含めないでください。'
          : 'Generate a more complete answer based on the information below. Maintain the original answer structure and incorporate information from the search results appropriately. Do not include any information if it is not available.',
        '\n\n--- ORIGINAL ANSWER ---\n',
        responseContent,
        '\n\n--- SEARCH RESULTS ---\n'
      ];
      
      // 検索結果の追加
      for (let i = 0; i < Math.min(responseData.results.length, 5); i++) {
        const result = responseData.results[i];
        promptParts.push(`Result ${i+1}: ${result.content || result.snippet || ''}\n\n`);
      }
      
      // 言語指定
      promptParts.push(`\n\n--- LANGUAGE ---\n${language}`);
      
      // システム指示
      const systemInstructions = [
        `You are a helpful assistant that enhances answers based on search results. 
        Your task is to improve the original answer while maintaining its structure and format. 
        Add relevant information from the search results to make the answer more complete and helpful.
        Do not hallucinate or add information not present in the search results.
        Always respond in the specified language: ${language}.`
      ];
      
      // Gemini APIを呼び出し
      const result = await GeminiIntegration.generateText(
        promptParts.join(''),
        {
          temperature: 0.2, // より決定論的な出力のために低い温度
          maxOutputTokens: 1500,
          systemInstructions: systemInstructions
        }
      );
      
      if (!result.success || !result.text) {
        return {
          success: false,
          content: responseContent,
          error: result.error || 'Unknown error'
        };
      }
      
      return {
        success: true,
        content: result.text
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.enhanceResponseWithGemini',
        error: error,
        severity: 'LOW',
        context: { response_length: responseContent.length },
        retry: false
      });
      
      // エラー時は元の応答を返す
      return {
        success: false,
        content: responseContent,
        error: error.message
      };
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
        case 'generateResponse':
          const result = await this.generateResponse(params.searchResults, params.query, params.options);
          return result.success;
          
        case 'enhanceResponseWithGemini':
          const enhanceResult = await this.enhanceResponseWithGemini(params.responseContent, params.responseData);
          return enhanceResult.success;
          
        default:
          console.warn(`Unknown operation for retry: ${operation}`);
          return false;
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'ResponseGenerator.retryOperation',
        error: error,
        severity: 'HIGH',
        context: options.context
      });
      
      return false;
    }
  }
}
