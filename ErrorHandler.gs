/**
 * RAG 2.0 エラーハンドラー
 * システム全体のエラー処理、ロギング、復旧を管理します
 */
class ErrorHandler {

  /**
   * エラーを処理します
   * @param {Object} options エラー処理オプション
   * @param {string} options.source エラーソース（発生場所）
   * @param {Error|string} options.error エラーオブジェクトまたはメッセージ
   * @param {string} options.severity エラー重大度 ('LOW', 'MEDIUM', 'HIGH', 'CRITICAL')
   * @param {Object} options.context 追加コンテキスト情報
   * @param {boolean} options.retry 再試行するかどうか
   * @return {boolean} エラーが正常に処理されたかどうか
   */
  static handleError(options) {
    try {
      const { source, error, severity = 'MEDIUM', context = {}, retry = false } = options;

      // エラーメッセージの標準化
      const errorMessage = error instanceof Error ? error.message : String(error);
      const stackTrace = error instanceof Error ? error.stack : null;

      // エラー情報の構築
      const errorInfo = {
        source: source,
        message: errorMessage,
        stack: stackTrace,
        severity: severity,
        timestamp: new Date().toISOString(),
        context: JSON.stringify(context)
      };

      // エラーのロギング
      this.logError(errorInfo);

      // 重大度に応じた処理
      switch (severity) {
        case 'CRITICAL':
          // クリティカルエラーは即座に管理者に通知
          this.notifyAdmins(errorInfo);
          break;

        case 'HIGH':
          // 高重大度エラーはログに記録し、可能であれば復旧を試みる
          if (retry) {
            return this.attemptRecovery(options);
          }
          break;

        case 'MEDIUM':
          // 中重大度エラーはログに記録
          // 特別な処理は不要
          break;

        case 'LOW':
          // 低重大度エラーは最小限のロギングのみ
          break;
      }

      return true;
    } catch (error) {
      // エラーハンドラ自体でエラーが発生した場合
      console.error('ErrorHandler failed:', error);
      return false;
    }
  }

  /**
   * エラーをログに記録します
   * @param {Object} errorInfo エラー情報
   */
  static logError(errorInfo) {
    try {
      // コンソールにログ出力
      console.error(`[${errorInfo.severity}] ${errorInfo.source}: ${errorInfo.message}`);

      // スプレッドシートにログを記録
      this.logToSpreadsheet(errorInfo);
    } catch (error) {
      console.error('Failed to log error:', error);
    }
  }

  /**
   * スプレッドシートにエラーログを記録します
   * @param {Object} errorInfo エラー情報
   */
  static logToSpreadsheet(errorInfo) {
    try {
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      const logSheet = ss.getSheetByName('Logs');

      if (!logSheet) {
        console.error('Logs sheet not found');
        return;
      }

      // ログデータの準備
      const logRow = [
        errorInfo.timestamp,
        errorInfo.severity,
        errorInfo.source,
        errorInfo.message,
        errorInfo.context,
        errorInfo.stack || 'N/A'
      ];

      // シートに追加
      logSheet.appendRow(logRow);

      // ログシートが大きくなりすぎないように古いログを削除
      const maxRows = 1000; // 保持する最大行数
      const currentRows = logSheet.getLastRow();

      if (currentRows > maxRows) {
        // 古いログを削除（先頭の行はヘッダーなので残す）
        const rowsToDelete = currentRows - maxRows;
        if (rowsToDelete > 0) {
          logSheet.deleteRows(2, rowsToDelete);
        }
      }
    } catch (error) {
      console.error('Failed to log to spreadsheet:', error);
    }
  }

  /**
   * 管理者にエラーを通知します
   * @param {Object} errorInfo エラー情報
   */
  static notifyAdmins(errorInfo) {
    try {
      const config = Config.getSystemConfig();
      const adminEmails = config.admin_emails;

      if (!adminEmails || adminEmails.length === 0) {
        console.error('No admin emails configured for notifications');
        return;
      }

      // メール通知の作成
      const subject = `[RAG System] ${errorInfo.severity} Error: ${errorInfo.source}`;
      const body = `
        Error Details:

        Source: ${errorInfo.source}
        Severity: ${errorInfo.severity}
        Timestamp: ${errorInfo.timestamp}
        Message: ${errorInfo.message}

        Context: ${errorInfo.context}

        Stack Trace:
        ${errorInfo.stack || 'Not available'}

        This is an automated message from the RAG System.
      `;

      // メール送信
      for (const email of adminEmails) {
        GmailApp.sendEmail(email, subject, body);
      }
    } catch (error) {
      console.error('Failed to notify admins:', error);
    }
  }

  /**
   * エラーからの復旧を試みます
   * @param {Object} options エラー処理オプション
   * @return {boolean} 復旧が成功したかどうか
   */
  static attemptRecovery(options) {
    try {
      const { source, error, context = {}, retryCount = 0 } = options;
      const config = Config.getSystemConfig();
      const maxRetries = config.max_retry_count || 3;

      // 最大再試行回数を超えていないか確認
      if (retryCount >= maxRetries) {
        console.warn(`Maximum retry count (${maxRetries}) exceeded for ${source}`);
        return false;
      }

      // エラータイプに基づいた復旧戦略
      const errorMessage = error instanceof Error ? error.message : String(error);
      const updatedOptions = { ...options, retryCount: retryCount + 1 };

      // 一時的なエラーの場合は再試行
      if (this.isTemporaryError(errorMessage)) {
        // 指数バックオフを使用した再試行
        const backoffTime = Math.pow(config.retry_backoff_base || 2, retryCount) * 1000;
        console.log(`Retrying ${source} in ${backoffTime}ms (attempt ${retryCount + 1}/${maxRetries})`);

        // 処理を遅延実行
        Utilities.sleep(backoffTime);

        // ソースに基づいて適切な回復関数を呼び出す
        return this.executeRecoveryFunction(source, updatedOptions);
      }

      // データ整合性エラーの場合はバックアップから復元
      if (this.isDataIntegrityError(errorMessage)) {
        console.log(`Attempting data recovery for ${source}`);
        return this.recoverFromBackup(source, context);
      }

      // その他のエラーの場合
      return false;
    } catch (error) {
      console.error('Recovery attempt failed:', error);
      return false;
    }
  }

  /**
   * 一時的なエラーかどうかを判断します
   * @param {string} errorMessage エラーメッセージ
   * @return {boolean} 一時的なエラーかどうか
   */
  static isTemporaryError(errorMessage) {
    const temporaryErrorPatterns = [
      /timeout/i,
      /rate limit/i,
      /quota/i,
      /temporarily unavailable/i,
      /too many requests/i,
      /server error/i,
      /network/i,
      /connection/i,
      /503/i,
      /429/i,
      /500/i
    ];

    return temporaryErrorPatterns.some(pattern => pattern.test(errorMessage));
  }

  /**
   * データ整合性エラーかどうかを判断します
   * @param {string} errorMessage エラーメッセージ
   * @return {boolean} データ整合性エラーかどうか
   */
  static isDataIntegrityError(errorMessage) {
    const dataErrorPatterns = [
      /corrupt/i,
      /integrity/i,
      /invalid format/i,
      /malformed/i,
      /inconsistent/i,
      /missing data/i,
      /not found/i
    ];

    return dataErrorPatterns.some(pattern => pattern.test(errorMessage));
  }

  /**
   * ソースに基づいて適切な回復関数を実行します
   * @param {string} source エラーソース
   * @param {Object} options エラー処理オプション
   * @return {boolean} 回復が成功したかどうか
   */
  static executeRecoveryFunction(source, options) {
    // ソースに基づいて適切な回復関数をマッピング
    const recoveryMap = {
      'SheetStorage': () => SheetStorage.retryOperation(options),
      'DocumentProcessor': () => DocumentProcessor.retryOperation(options),
      'EmbeddingManager': () => EmbeddingManager.retryOperation(options),
      'GeminiIntegration': () => GeminiIntegration.retryOperation(options),
      // 他のモジュールの回復関数をマッピング
    };

    // ソースの先頭部分に基づいてマッピング（例：'SheetStorage.getChunks' → 'SheetStorage'）
    const moduleSource = source.split('.')[0];
    const recoveryFunction = recoveryMap[moduleSource];

    if (typeof recoveryFunction === 'function') {
      return recoveryFunction();
    }

    // 該当する回復関数がない場合
    console.warn(`No recovery function defined for source: ${source}`);
    return false;
  }

  /**
   * バックアップからデータを復元します
   * @param {string} source エラーソース
   * @param {Object} context コンテキスト情報
   * @return {boolean} 復元が成功したかどうか
   */
  static recoverFromBackup(source, context) {
    try {
      console.log(`Attempting to recover ${source} from backup`);

      // モジュールに基づいた復元戦略
      if (source.startsWith('SheetStorage')) {
        return this.recoverSheetFromBackup(context);
      }

      return false;
    } catch (error) {
      console.error('Recovery from backup failed:', error);
      return false;
    }
  }

  /**
   * スプレッドシートをバックアップから復元します
   * @param {Object} context コンテキスト情報
   * @return {boolean} 復元が成功したかどうか
   */
  static recoverSheetFromBackup(context) {
    try {
      const { sheetName } = context;
      if (!sheetName) {
        console.error('Sheet name not provided in context for backup recovery');
        return false;
      }

      // 最新のバックアップを探す
      const config = Config.getSystemConfig();
      const driveApp = DriveApp;
      let backupFolderId = config.backup_folder_id;
      
      // バックアップフォルダIDが設定されていない場合はルートフォルダからBackupsを探す
      if (!backupFolderId) {
        const rootFolderId = Config.getRootFolderId();
        const rootFolder = driveApp.getFolderById(rootFolderId);
        const backupFolders = rootFolder.getFoldersByName('Backups');
        if (backupFolders.hasNext()) {
          backupFolderId = backupFolders.next().getId();
        } else {
          console.error('Backup folder not found');
          return false;
        }
      }

      const backupFolder = driveApp.getFolderById(backupFolderId);
      const backupFiles = backupFolder.getFilesByName(`${sheetName}_backup`);
      
      if (!backupFiles.hasNext()) {
        console.error(`No backup found for sheet: ${sheetName}`);
        return false;
      }

      // 最新のバックアップを取得
      let latestBackup = null;
      let latestDate = 0;

      while (backupFiles.hasNext()) {
        const file = backupFiles.next();
        const fileDate = file.getLastUpdated().getTime();

        if (fileDate > latestDate) {
          latestBackup = file;
          latestDate = fileDate;
        }
      }

      if (!latestBackup) {
        console.error('Failed to identify latest backup');
        return false;
      }

      // バックアップからデータを読み込む
      const backupData = latestBackup.getBlob().getDataAsString();
      const backupJson = JSON.parse(backupData);

      // 対象のスプレッドシートを取得
      const ss = SpreadsheetApp.openById(Config.getDatabaseId());
      let sheet = ss.getSheetByName(sheetName);

      // シートが存在しない場合は作成
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        // 既存データをクリア
        sheet.clear();
      }

      // バックアップデータを書き込む
      if (backupJson.headers && backupJson.headers.length > 0) {
        sheet.appendRow(backupJson.headers);
      }

      if (backupJson.data && backupJson.data.length > 0) {
        sheet.getRange(2, 1, backupJson.data.length, backupJson.data[0].length)
          .setValues(backupJson.data);
      }

      console.log(`Successfully recovered sheet ${sheetName} from backup`);
      return true;
    } catch (error) {
      console.error('Sheet recovery from backup failed:', error);
      return false;
    }
  }
  
  /**
   * 適切なフォールバックメッセージを作成します
   * @param {string} operation 失敗した操作
   * @return {string} フォールバックメッセージ
   */
  static createFallbackMessage(operation) {
    const fallbackMessages = {
      'search': '申し訳ありませんが、検索処理中にエラーが発生しました。しばらく経ってからもう一度お試しください。',
      'embedding': '埋め込み処理中にエラーが発生しました。システム管理者に連絡してください。',
      'document_processing': 'ドキュメント処理中にエラーが発生しました。ファイル形式を確認して再度お試しください。',
      'response_generation': '応答生成中にエラーが発生しました。別のクエリで試すか、しばらく経ってからもう一度お試しください。',
      'api_call': 'APIサービスへの接続中にエラーが発生しました。インターネット接続を確認して再度お試しください。',
      'default': 'エラーが発生しました。しばらく経ってからもう一度お試しください。'
    };
    
    return fallbackMessages[operation] || fallbackMessages.default;
  }
  
  /**
   * エラー発生時のユーザー向けメッセージを作成します
   * @param {Object} errorInfo エラー情報
   * @param {string} language 言語コード
   * @return {Object} ユーザー向けエラーメッセージと対応策
   */
  static createUserMessage(errorInfo, language = 'ja') {
    const errorType = this.classifyErrorType(errorInfo.message);
    
    // 日本語メッセージ
    const jaMessages = {
      'temporary': {
        message: '一時的なエラーが発生しました。',
        action: 'しばらく経ってからもう一度お試しください。'
      },
      'quota': {
        message: 'サービスの利用制限に達しました。',
        action: 'しばらく待ってから再度お試しいただくか、システム管理者にお問い合わせください。'
      },
      'auth': {
        message: '認証エラーが発生しました。',
        action: 'ログインし直すか、APIキーを確認してください。'
      },
      'data': {
        message: 'データ処理中にエラーが発生しました。',
        action: '入力データを確認して再度お試しください。'
      },
      'system': {
        message: 'システムエラーが発生しました。',
        action: 'システム管理者に連絡してください。エラーID: ' + errorInfo.timestamp
      },
      'unknown': {
        message: '予期しないエラーが発生しました。',
        action: 'もう一度お試しいただくか、別の操作を行ってください。'
      }
    };
    
    // 英語メッセージ
    const enMessages = {
      'temporary': {
        message: 'A temporary error occurred.',
        action: 'Please try again after a moment.'
      },
      'quota': {
        message: 'Service quota limit reached.',
        action: 'Please wait and try again later or contact your system administrator.'
      },
      'auth': {
        message: 'Authentication error occurred.',
        action: 'Please try logging in again or check your API key.'
      },
      'data': {
        message: 'Error occurred while processing data.',
        action: 'Please check your input data and try again.'
      },
      'system': {
        message: 'System error occurred.',
        action: 'Please contact your system administrator. Error ID: ' + errorInfo.timestamp
      },
      'unknown': {
        message: 'An unexpected error occurred.',
        action: 'Please try again or perform a different operation.'
      }
    };
    
    const messages = language === 'ja' ? jaMessages : enMessages;
    const errorInfo = messages[errorType] || messages.unknown;
    
    return {
      message: errorInfo.message,
      action: errorInfo.action,
      errorType: errorType,
      severity: errorInfo.severity
    };
  }
  
  /**
   * エラーメッセージからエラータイプを分類します
   * @param {string} errorMessage エラーメッセージ
   * @return {string} エラータイプ
   */
  static classifyErrorType(errorMessage) {
    if (/timeout|retry|unavailable|503|500/i.test(errorMessage)) {
      return 'temporary';
    } else if (/rate limit|quota|too many requests|429/i.test(errorMessage)) {
      return 'quota';
    } else if (/auth|unauthorized|permission|access denied|401|403/i.test(errorMessage)) {
      return 'auth';
    } else if (/data|format|parse|syntax|value|not found|404/i.test(errorMessage)) {
      return 'data';
    } else if (/system|internal|memory|crash|exception/i.test(errorMessage)) {
      return 'system';
    } else {
      return 'unknown';
    }
  }
}
