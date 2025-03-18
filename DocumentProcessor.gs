/**
 * RAG 2.0 ドキュメント処理モジュール
 * 共有ドライブからのドキュメント取得と処理を管理します
 */
class DocumentProcessor {

  /**
   * ドキュメントを取得して処理します
   * @param {string} fileId ファイルID
   * @param {Object} options 処理オプション
   * @return {Object} 処理結果
   */
  static processDocument(fileId, options = {}) {
    try {
      // ファイルを取得
      const file = DriveApp.getFileById(fileId);
      if (!file) {
        throw new Error(`ファイル ${fileId} が見つかりません`);
      }

      // ファイル情報を取得
      const fileName = file.getName();
      const fileType = file.getMimeType();
      const fileDate = file.getLastUpdated();
      const filePath = this.getFilePath(file);

      // ファイルタイプに基づいてテキスト抽出
      const textContent = this.extractText(file, fileType);
      if (!textContent) {
        throw new Error(`ファイル ${fileName} からテキストを抽出できませんでした`);
      }

      // メタデータを構築
      const metadata = {
        title: fileName,
        path: filePath,
        format: this.getFormatFromMimeType(fileType),
        last_updated: fileDate.toISOString(),
        source: "google_drive",
        file_id: fileId
      };

      // ファイル名から言語とカテゴリを判定
      const { language, category } = this.detectLanguageAndCategory(fileName, filePath);
      metadata.language = language;

      // ドキュメントオブジェクトを構築
      const document = {
        id: `doc_${fileId}`,
        title: fileName,
        content: textContent,
        category: category,
        language: language,
        path: filePath,
        format: this.getFormatFromMimeType(fileType),
        metadata: metadata
      };

      // ドキュメントをストレージに保存
      const saveResult = SheetStorage.saveDocument(document);

      return {
        success: saveResult.success,
        document_id: document.id,
        file_id: fileId,
        title: fileName,
        chunks_count: saveResult.chunks_count || 0,
        language: language,
        category: category
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.processDocument',
        error: error,
        severity: 'MEDIUM',
        context: { file_id: fileId }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * フォルダ内のすべてのドキュメントを処理します
   * @param {string} folderId フォルダID
   * @param {Object} options 処理オプション
   * @param {boolean} options.recursive サブフォルダも処理するかどうか
   * @param {Array<string>} options.fileTypes 処理対象のファイル形式
   * @param {boolean} options.updateOnly 更新されたドキュメントのみ処理するかどうか
   * @return {Object} 処理結果
   */
  static processFolderDocuments(folderId, options = {}) {
    try {
      // デフォルトオプション
      const defaultOptions = {
        recursive: true,
        fileTypes: ['application/pdf', 'application/vnd.google-apps.document', 'text/html', 'text/plain'],
        updateOnly: true
      };

      // オプションをマージ
      const mergedOptions = { ...defaultOptions, ...options };

      // フォルダを取得
      const folder = DriveApp.getFolderById(folderId);
      if (!folder) {
        throw new Error(`フォルダ ${folderId} が見つかりません`);
      }

      // 結果を初期化
      const results = {
        success: true,
        processedCount: 0,
        skippedCount: 0,
        errorCount: 0,
        errors: [],
        subFolders: []
      };

      // フォルダ内のファイルを処理
      this.processFilesInFolder(folder, results, mergedOptions);

      // サブフォルダを再帰的に処理
      if (mergedOptions.recursive) {
        const subFolders = folder.getFolders();
        while (subFolders.hasNext()) {
          const subFolder = subFolders.next();
          const subFolderId = subFolder.getId();
          const subFolderName = subFolder.getName();
          
          // サブフォルダの結果を記録
          const subFolderResult = {
            id: subFolderId,
            name: subFolderName,
            processedCount: 0,
            skippedCount: 0,
            errorCount: 0
          };
          results.subFolders.push(subFolderResult);
          
          // サブフォルダの処理結果を更新
          this.processFilesInFolder(subFolder, subFolderResult, mergedOptions);
          
          // 親の合計を更新
          results.processedCount += subFolderResult.processedCount;
          results.skippedCount += subFolderResult.skippedCount;
          results.errorCount += subFolderResult.errorCount;
        }
      }

      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.processFolderDocuments',
        error: error,
        severity: 'HIGH',
        context: { folder_id: folderId }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * フォルダ内のファイルを処理します
   * @param {Folder} folder フォルダオブジェクト
   * @param {Object} results 結果オブジェクト
   * @param {Object} options 処理オプション
   */
  static processFilesInFolder(folder, results, options) {
    // フォルダ内のファイルを取得
    const files = folder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();
      const fileName = file.getName();
      const fileType = file.getMimeType();
      const lastUpdated = file.getLastUpdated();
      
      // ファイルタイプがオプションで指定されたものに含まれるか確認
      if (!options.fileTypes.includes(fileType) && !options.fileTypes.includes('*')) {
        results.skippedCount++;
        continue;
      }
      
      // 更新のみモードの場合、既存のドキュメントを確認
      if (options.updateOnly) {
        const existingDoc = SheetStorage.getDocumentById(`doc_${fileId}`);
        if (existingDoc) {
          // 最終更新日時を比較
          const existingLastUpdated = new Date(existingDoc.last_updated || 0);
          if (existingLastUpdated >= lastUpdated) {
            // 更新が必要ない場合はスキップ
            results.skippedCount++;
            continue;
          }
        }
      }
      
      // ドキュメントを処理
      try {
        const processResult = this.processDocument(fileId, options);
        if (processResult.success) {
          results.processedCount++;
        } else {
          results.errorCount++;
          results.errors.push({
            file_id: fileId,
            file_name: fileName,
            error: processResult.error
          });
        }
      } catch (error) {
        results.errorCount++;
        results.errors.push({
          file_id: fileId,
          file_name: fileName,
          error: error.message
        });
      }
    }
  }

  /**
   * ヘルプページのペアを処理します（日英言語対応）
   * @param {string} folderId ヘルプページフォルダID
   * @return {Object} 処理結果
   */
  static processHelpPagePairs(folderId) {
    try {
      // フォルダを取得
      const folder = DriveApp.getFolderById(folderId);
      if (!folder) {
        throw new Error(`フォルダ ${folderId} が見つかりません`);
      }

      // 結果を初期化
      const results = {
        success: true,
        pairsProcessed: 0,
        pairsSkipped: 0,
        individualProcessed: 0,
        errors: []
      };

      // フォルダ内のファイルを収集
      const files = {};
      const fileIterator = folder.getFiles();
      
      while (fileIterator.hasNext()) {
        const file = fileIterator.next();
        const fileName = file.getName();
        files[fileName] = file;
      }

      // ペアを見つけて処理
      const processedFiles = new Set();
      
      for (const fileName in files) {
        // すでに処理済みならスキップ
        if (processedFiles.has(fileName)) {
          continue;
        }
        
        // 日本語版と英語版のペアを検索
        const baseName = this.getBaseNameFromHelpPage(fileName);
        if (!baseName) {
          continue;
        }
        
        const jaName = `${baseName}_ja.mhtml`;
        const enName = `${baseName}_en.mhtml`;
        
        const jaFile = files[jaName];
        const enFile = files[enName];
        
        if (jaFile && enFile) {
          // ペアを処理
          const pairResult = this.processHelpPagePair(jaFile, enFile);
          
          if (pairResult.success) {
            results.pairsProcessed++;
          } else {
            results.errors.push({
              base_name: baseName,
              error: pairResult.error
            });
          }
          
          // 処理済みとしてマーク
          processedFiles.add(jaName);
          processedFiles.add(enName);
        } else {
          // 個別のファイルを処理
          const fileToProcess = jaFile || enFile;
          if (fileToProcess) {
            const processResult = this.processDocument(fileToProcess.getId());
            
            if (processResult.success) {
              results.individualProcessed++;
            } else {
              results.errors.push({
                file_name: fileToProcess.getName(),
                error: processResult.error
              });
            }
            
            // 処理済みとしてマーク
            processedFiles.add(fileToProcess.getName());
          }
        }
      }

      // 結果を返す
      return results;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.processHelpPagePairs',
        error: error,
        severity: 'HIGH',
        context: { folder_id: folderId }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * ヘルプページのペアを処理します
   * @param {File} jaFile 日本語版ファイル
   * @param {File} enFile 英語版ファイル
   * @return {Object} 処理結果
   */
  static processHelpPagePair(jaFile, enFile) {
    try {
      // ファイル情報を取得
      const jaFileId = jaFile.getId();
      const enFileId = enFile.getId();
      const baseName = this.getBaseNameFromHelpPage(jaFile.getName());
      
      // ファイルからテキストを抽出
      const jaText = this.extractText(jaFile, jaFile.getMimeType());
      const enText = this.extractText(enFile, enFile.getMimeType());
      
      if (!jaText || !enText) {
        throw new Error('テキスト抽出に失敗しました');
      }
      
      // 日本語版ドキュメントを作成
      const jaDocument = {
        id: `doc_${jaFileId}`,
        title: jaFile.getName(),
        content: jaText,
        category: 'Help_Pages',
        language: 'ja',
        path: this.getFilePath(jaFile),
        format: 'mhtml',
        metadata: {
          title: jaFile.getName(),
          path: this.getFilePath(jaFile),
          format: 'mhtml',
          last_updated: jaFile.getLastUpdated().toISOString(),
          source: "google_drive",
          file_id: jaFileId,
          language: 'ja',
          help_base_name: baseName,
          paired_document_id: `doc_${enFileId}`
        }
      };
      
      // 英語版ドキュメントを作成
      const enDocument = {
        id: `doc_${enFileId}`,
        title: enFile.getName(),
        content: enText,
        category: 'Help_Pages',
        language: 'en',
        path: this.getFilePath(enFile),
        format: 'mhtml',
        metadata: {
          title: enFile.getName(),
          path: this.getFilePath(enFile),
          format: 'mhtml',
          last_updated: enFile.getLastUpdated().toISOString(),
          source: "google_drive",
          file_id: enFileId,
          language: 'en',
          help_base_name: baseName,
          paired_document_id: `doc_${jaFileId}`
        }
      };
      
      // ドキュメントをストレージに保存
      const jaSaveResult = SheetStorage.saveDocument(jaDocument);
      const enSaveResult = SheetStorage.saveDocument(enDocument);
      
      // ヘルプペア情報を保存
      this.saveHelpPair(baseName, jaDocument.id, enDocument.id);
      
      return {
        success: jaSaveResult.success && enSaveResult.success,
        base_name: baseName,
        ja_document_id: jaDocument.id,
        en_document_id: enDocument.id,
        ja_chunks_count: jaSaveResult.chunks_count || 0,
        en_chunks_count: enSaveResult.chunks_count || 0
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.processHelpPagePair',
        error: error,
        severity: 'MEDIUM',
        context: { 
          ja_file_id: jaFile.getId(),
          en_file_id: enFile.getId()
        }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * ヘルプページのベース名を取得します
   * @param {string} fileName ファイル名
   * @return {string|null} ベース名
   */
  static getBaseNameFromHelpPage(fileName) {
    // ファイル名から言語サフィックスを削除
    const match = fileName.match(/(.+)_(ja|en)\.mhtml$/);
    if (match && match[1]) {
      return match[1];
    }
    return null;
  }

  /**
   * ヘルプページのペア情報を保存します
   * @param {string} baseName ベース名
   * @param {string} jaDocumentId 日本語版ドキュメントID
   * @param {string} enDocumentId 英語版ドキュメントID
   * @return {Object} 保存結果
   */
  static saveHelpPair(baseName, jaDocumentId, enDocumentId) {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const helpPairsSheet = ss.getSheetByName('Help_Pairs');
      
      if (!helpPairsSheet) {
        throw new Error('Help_Pairs シートが見つかりません');
      }
      
      // ペアIDを生成
      const pairId = `pair_${Utilities.generateUniqueId()}`;
      
      // データを準備
      const data = [
        pairId,
        baseName,
        jaDocumentId,
        enDocumentId,
        'Help_Pages',
        new Date().toISOString(),
        'active'
      ];
      
      // データを保存
      helpPairsSheet.appendRow(data);
      
      return {
        success: true,
        pair_id: pairId
      };
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.saveHelpPair',
        error: error,
        severity: 'MEDIUM',
        context: { 
          base_name: baseName,
          ja_document_id: jaDocumentId,
          en_document_id: enDocumentId
        }
      });

      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * ファイルからテキストを抽出します
   * @param {File} file ファイルオブジェクト
   * @param {string} mimeType MIME タイプ
   * @return {string} 抽出されたテキスト
   */
  static extractText(file, mimeType) {
    try {
      // MIME タイプに応じて適切な抽出方法を使用
      switch (mimeType) {
        case 'application/pdf':
          return this.extractPdfText(file);
        
        case 'application/vnd.google-apps.document':
          return this.extractGoogleDocText(file);
        
        case 'application/vnd.google-apps.spreadsheet':
          return this.extractGoogleSheetText(file);
        
        case 'application/vnd.google-apps.presentation':
          return this.extractGoogleSlideText(file);
        
        case 'text/html':
        case 'application/xhtml+xml':
        case 'message/rfc822': // MHTML files
          return this.extractHtmlText(file);
        
        case 'text/plain':
          return this.extractPlainText(file);
        
        default:
          // サポートされていない形式
          throw new Error(`サポートされていないファイル形式です: ${mimeType}`);
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.extractText',
        error: error,
        severity: 'MEDIUM',
        context: { 
          file_id: file.getId(),
          mime_type: mimeType
        }
      });

      return null;
    }
  }

  /**
   * PDF ファイルからテキストを抽出します
   * @param {File} file ファイルオブジェクト
   * @return {string} 抽出されたテキスト
   */
  static extractPdfText(file) {
    try {
      // PDF を Google ドキュメントに変換
      const blob = file.getBlob();
      const resource = {
        title: file.getName(),
        mimeType: 'application/vnd.google-apps.document'
      };
      
      // ドライブに一時的なドキュメントを作成
      const tempDocFile = Drive.Files.insert(resource, blob);
      const tempDoc = DocumentApp.openById(tempDocFile.id);
      
      // テキストを抽出
      const text = tempDoc.getBody().getText();
      
      // 一時ファイルを削除
      Drive.Files.remove(tempDocFile.id);
      
      return text;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.extractPdfText',
        error: error,
        severity: 'MEDIUM',
        context: { file_id: file.getId() }
      });

      // 代替方法: PDF データを直接取得して解析
      try {
        const blob = file.getBlob();
        // ここで PDF 本文を解析する代替方法を実装できますが、
        // Google Apps Script では PDF を直接解析する標準ライブラリがないため、
        // 外部サービスの利用やシンプルなテキスト抽出に限られます。
        return "PDF の内容を直接抽出できませんでした。Google ドキュメントへの変換で問題が発生しました。";
      } catch (fallbackError) {
        ErrorHandler.handleError({
          source: 'DocumentProcessor.extractPdfText (fallback)',
          error: fallbackError,
          severity: 'HIGH',
          context: { file_id: file.getId() }
        });
        return null;
      }
    }
  }

  /**
   * Google ドキュメントからテキストを抽出します
   * @param {File} file ファイルオブジェクト
   * @return {string} 抽出されたテキスト
   */
  static extractGoogleDocText(file) {
    try {
      const doc = DocumentApp.openById(file.getId());
      return doc.getBody().getText();
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.extractGoogleDocText',
        error: error,
        severity: 'MEDIUM',
        context: { file_id: file.getId() }
      });
      
      // 代替方法: ドキュメントをプレーンテキストとしてエクスポート
      try {
        const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${file.getId()}&exportFormat=txt`;
        const options = {
          headers: {
            'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`
          },
          muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(url, options);
        return response.getContentText();
      } catch (fallbackError) {
        ErrorHandler.handleError({
          source: 'DocumentProcessor.extractGoogleDocText (fallback)',
          error: fallbackError,
          severity: 'HIGH',
          context: { file_id: file.getId() }
        });
        return null;
      }
    }
  }

  /**
   * Google スプレッドシートからテキストを抽出します
   * @param {File} file ファイルオブジェクト
   * @return {string} 抽出されたテキスト
   */
  static extractGoogleSheetText(file) {
    try {
      const ss = SpreadsheetApp.openById(file.getId());
      const sheets = ss.getSheets();
      let text = `Spreadsheet: ${file.getName()}\n\n`;
      
      for (const sheet of sheets) {
        const sheetName = sheet.getName();
        text += `Sheet: ${sheetName}\n`;
        
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        
        if (values.length > 0) {
          for (const row of values) {
            text += row.join('\t') + '\n';
          }
        } else {
          text += '(Empty Sheet)\n';
        }
        
        text += '\n';
      }
      
      return text;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.extractGoogleSheetText',
        error: error,
        severity: 'MEDIUM',
        context: { file_id: file.getId() }
      });
      return null;
    }
  }

  /**
   * Google スライドからテキストを抽出します
   * @param {File} file ファイルオブジェクト
   * @return {string} 抽出されたテキスト
   */
  static extractGoogleSlideText(file) {
    try {
      const presentation = SlidesApp.openById(file.getId());
      const slides = presentation.getSlides();
      let text = `Presentation: ${file.getName()}\n\n`;
      
      for (let i = 0; i < slides.length; i++) {
        const slide = slides[i];
        text += `Slide ${i + 1}:\n`;
        
        // スライドタイトルを抽出
        const pageElements = slide.getPageElements();
        for (const element of pageElements) {
          if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
            const shape = element.asShape();
            if (shape.getText()) {
              text += shape.getText().asString() + '\n';
            }
          }
        }
        
        text += '\n';
      }
      
      return text;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.extractGoogleSlideText',
        error: error,
        severity: 'MEDIUM',
        context: { file_id: file.getId() }
      });
      
      // 代替方法: スライドをテキストとしてエクスポート
      try {
        const url = `https://docs.google.com/feeds/download/presentations/export/Export?id=${file.getId()}&exportFormat=txt`;
        const options = {
          headers: {
            'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`
          },
          muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(url, options);
        return response.getContentText();
      } catch (fallbackError) {
        ErrorHandler.handleError({
          source: 'DocumentProcessor.extractGoogleSlideText (fallback)',
          error: fallbackError,
          severity: 'HIGH',
          context: { file_id: file.getId() }
        });
        return null;
      }
    }
  }

  /**
   * HTML/MHTML ファイルからテキストを抽出します
   * @param {File} file ファイルオブジェクト
   * @return {string} 抽出されたテキスト
   */
  static extractHtmlText(file) {
    try {
      const content = file.getBlob().getDataAsString();
      
      // 簡易的な HTML パーサー
      // 実際のプロジェクトではより堅牢なライブラリを使用することをお勧めします
      let text = content
        // HTML タグを削除
        .replace(/<[^>]*>/g, ' ')
        // スクリプトとスタイルの内容を削除
        .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
        .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
        // HTML エンティティをデコード
        .replace(/&nbsp;/g, ' ')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&')
        .replace(/&quot;/g, '"')
        // 余分な空白を圧縮
        .replace(/\s+/g, ' ')
        .trim();
      
      return text;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.extractHtmlText',
        error: error,
        severity: 'MEDIUM',
        context: { file_id: file.getId() }
      });
      return null;
    }
  }

  /**
   * プレーンテキストファイルからテキストを抽出します
   * @param {File} file ファイルオブジェクト
   * @return {string} 抽出されたテキスト
   */
  static extractPlainText(file) {
    try {
      return file.getBlob().getDataAsString();
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.extractPlainText',
        error: error,
        severity: 'MEDIUM',
        context: { file_id: file.getId() }
      });
      return null;
    }
  }

  /**
   * MIME タイプから内部形式名を取得します
   * @param {string} mimeType MIME タイプ
   * @return {string} 内部形式名
   */
  static getFormatFromMimeType(mimeType) {
    const mimeToFormat = {
      'application/pdf': 'pdf',
      'application/vnd.google-apps.document': 'doc',
      'application/vnd.google-apps.spreadsheet': 'sheet',
      'application/vnd.google-apps.presentation': 'slide',
      'text/html': 'html',
      'application/xhtml+xml': 'html',
      'message/rfc822': 'mhtml',
      'text/plain': 'text'
    };
    
    return mimeToFormat[mimeType] || 'unknown';
  }

  /**
   * ファイルのパスを取得します
   * @param {File} file ファイルオブジェクト
   * @return {string} ファイルパス
   */
  static getFilePath(file) {
    try {
      const fileName = file.getName();
      let path = fileName;
      
      // 親フォルダを取得
      const parents = file.getParents();
      if (parents.hasNext()) {
        const parentFolder = parents.next();
        const parentPath = this.getFolderPath(parentFolder);
        path = parentPath + '/' + fileName;
      }
      
      return path;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.getFilePath',
        error: error,
        severity: 'LOW',
        context: { file_id: file.getId() }
      });
      
      return file.getName();
    }
  }

  /**
   * フォルダのパスを取得します
   * @param {Folder} folder フォルダオブジェクト
   * @return {string} フォルダパス
   */
  static getFolderPath(folder) {
    try {
      const folderName = folder.getName();
      
      // 親フォルダを取得
      const parents = folder.getParents();
      if (parents.hasNext()) {
        const parentFolder = parents.next();
        const parentPath = this.getFolderPath(parentFolder);
        return parentPath + '/' + folderName;
      }
      
      return folderName;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.getFolderPath',
        error: error,
        severity: 'LOW',
        context: { folder_id: folder.getId() }
      });
      
      return folder.getName();
    }
  }

  /**
   * ファイル名とパスから言語とカテゴリを検出します
   * @param {string} fileName ファイル名
   * @param {string} filePath ファイルパス
   * @return {Object} 言語とカテゴリ
   */
  static detectLanguageAndCategory(fileName, filePath) {
    // 言語検出
    let language = 'en'; // デフォルト言語
    
    // ファイル名から言語を検出
    if (fileName.includes('_ja.') || fileName.endsWith('_ja')) {
      language = 'ja';
    } else if (fileName.includes('_en.') || fileName.endsWith('_en')) {
      language = 'en';
    }
    
    // カテゴリ検出
    let category = 'general'; // デフォルトカテゴリ
    
    // パスからカテゴリを検出
    const categoryMap = {
      'Help_Pages': ['help', 'help_pages', 'manual'],
      'Search': ['search', 'keyword', 'query'],
      'Mobile': ['mobile', 'app', 'android', 'ios'],
      'Shopping': ['shopping', 'ecommerce', 'product'],
      'Display': ['display', 'banner', 'image'],
      'Video': ['video', 'youtube', 'motion'],
      'M&A': ['measurement', 'analytics', 'conversion'],
      'Billing': ['billing', 'payment', 'invoice'],
      'Policy': ['policy', 'compliance', 'rule']
    };
    
    // パスを小文字に変換
    const lowerPath = filePath.toLowerCase();
    
    // カテゴリマッピングを検索
    for (const [cat, keywords] of Object.entries(categoryMap)) {
      for (const keyword of keywords) {
        if (lowerPath.includes(keyword)) {
          category = cat;
          break;
        }
      }
      
      if (category !== 'general') {
        break;
      }
    }
    
    // 特定のフォルダパターンでカテゴリを優先的に設定
    if (lowerPath.includes('/help_pages/')) {
      category = 'Help_Pages';
    } else if (lowerPath.includes('/search/')) {
      category = 'Search';
    } else if (lowerPath.includes('/mobile/')) {
      category = 'Mobile';
    } else if (lowerPath.includes('/shopping/')) {
      category = 'Shopping';
    } else if (lowerPath.includes('/display/')) {
      category = 'Display';
    } else if (lowerPath.includes('/video/')) {
      category = 'Video';
    } else if (lowerPath.includes('/m&a/') || lowerPath.includes('/measurement/')) {
      category = 'M&A';
    } else if (lowerPath.includes('/billing/')) {
      category = 'Billing';
    } else if (lowerPath.includes('/policy/')) {
      category = 'Policy';
    }
    
    return { language, category };
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
        case 'processDocument':
          return this.processDocument(params.fileId, params.options).success;
          
        case 'extractText':
          return !!this.extractText(params.file, params.mimeType);
          
        default:
          console.warn(`Unknown operation for retry: ${operation}`);
          return false;
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'DocumentProcessor.retryOperation',
        error: error,
        severity: 'HIGH',
        context: options.context
      });
      
      return false;
    }
  }
}
