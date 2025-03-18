/**
 * RAG 2.0 キャッシュマネージャー
 * 多層キャッシュを管理し、パフォーマンスを最適化します
 */
class CacheManager {

  /**
   * キャッシュからデータを取得します
   * @param {string} key キャッシュキー
   * @param {string} scope キャッシュスコープ ('user', 'script', 'document')
   * @return {string|null} キャッシュされたデータ（存在しない場合はnull）
   */
  static get(key, scope = 'script') {
    try {
      // まずメモリキャッシュを確認
      const memCacheKey = `mem_${scope}_${key}`;
      if (this.memoryCache[memCacheKey]) {
        const item = this.memoryCache[memCacheKey];
        if (item.expiresAt > Date.now()) {
          return item.value;
        } else {
          // 期限切れならメモリキャッシュから削除
          delete this.memoryCache[memCacheKey];
        }
      }

      // CacheServiceを確認
      const cache = this.getCacheByScope(scope);
      const data = cache.get(key);

      if (data) {
        // メモリキャッシュに追加（短い有効期間で）
        this.memoryCache[memCacheKey] = {
          value: data,
          expiresAt: Date.now() + 60000 // 1分間
        };
        return data;
      }

      // PropertiesServiceを確認（長期キャッシュ）
      const props = this.getPropertiesByScope(scope);
      const propKey = `cache_${key}`;
      const propData = props.getProperty(propKey);

      if (propData) {
        try {
          const parsedData = JSON.parse(propData);
          const { value, expiresAt } = parsedData;

          // 有効期限をチェック
          if (!expiresAt || expiresAt > Date.now()) {
            // CacheServiceにも追加（今後の高速アクセスのため）
            cache.put(key, value, 21600); // 6時間

            // メモリキャッシュに追加
            this.memoryCache[memCacheKey] = {
              value: value,
              expiresAt: Date.now() + 60000 // 1分間
            };

            return value;
          } else {
            // 期限切れなら削除
            props.deleteProperty(propKey);
          }
        } catch (e) {
          // JSON解析エラー、破損データとみなして削除
          props.deleteProperty(propKey);
        }
      }

      return null;
    } catch (error) {
      ErrorHandler.handleError({
        source: 'CacheManager.get',
        error: error,
        severity: 'LOW'
      });
      return null;
    }
  }

  /**
   * データをキャッシュに保存します
   * @param {string} key キャッシュキー
   * @param {string} value 保存する値
   * @param {number} ttl 有効期間（秒）
   * @param {string} scope キャッシュスコープ ('user', 'script', 'document')
   */
  static set(key, value, ttl = 3600, scope = 'script') {
    try {
      // TTLの上限を設定（7日）
      const maxTtl = 604800;
      ttl = Math.min(ttl, maxTtl);

      // メモリキャッシュに保存
      const memCacheKey = `mem_${scope}_${key}`;
      this.memoryCache[memCacheKey] = {
        value: value,
        expiresAt: Date.now() + (ttl * 1000)
      };

      // CacheServiceに保存（最大6時間）
      const cache = this.getCacheByScope(scope);
      const cacheTtl = Math.min(ttl, 21600);
      cache.put(key, value, cacheTtl);

      // 長期保存が必要な場合はPropertiesServiceにも保存
      if (ttl > 21600) {
        const props = this.getPropertiesByScope(scope);
        const propKey = `cache_${key}`;
        const propData = JSON.stringify({
          value: value,
          expiresAt: Date.now() + (ttl * 1000)
        });
        props.setProperty(propKey, propData);
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'CacheManager.set',
        error: error,
        severity: 'LOW'
      });
    }
  }

  /**
   * キャッシュからデータを削除します
   * @param {string} key キャッシュキー
   * @param {string} scope キャッシュスコープ ('user', 'script', 'document')
   */
  static remove(key, scope = 'script') {
    try {
      // メモリキャッシュから削除
      const memCacheKey = `mem_${scope}_${key}`;
      delete this.memoryCache[memCacheKey];

      // CacheServiceから削除
      const cache = this.getCacheByScope(scope);
      cache.remove(key);

      // PropertiesServiceから削除
      const props = this.getPropertiesByScope(scope);
      const propKey = `cache_${key}`;
      props.deleteProperty(propKey);
    } catch (error) {
      ErrorHandler.handleError({
        source: 'CacheManager.remove',
        error: error,
        severity: 'LOW'
      });
    }
  }

  /**
   * 指定したプレフィックスで始まるすべてのキャッシュを削除します
   * @param {string} prefix キープレフィックス
   * @param {string} scope キャッシュスコープ ('user', 'script', 'document')
   */
  static removeByPrefix(prefix, scope = 'script') {
    try {
      // メモリキャッシュから削除
      const memPrefix = `mem_${scope}_${prefix}`;
      for (const key in this.memoryCache) {
        if (key.startsWith(memPrefix)) {
          delete this.memoryCache[key];
        }
      }

      // PropertiesServiceから削除
      const props = this.getPropertiesByScope(scope);
      const allKeys = props.getKeys();
      const propPrefix = `cache_${prefix}`;

      const keysToDelete = allKeys.filter(k => k.startsWith(propPrefix));
      for (const key of keysToDelete) {
        props.deleteProperty(key);
      }

      // CacheServiceは直接プレフィックスで削除する機能がないため、
      // 個別のキーが分かっている場合のみ対応
    } catch (error) {
      ErrorHandler.handleError({
        source: 'CacheManager.removeByPrefix',
        error: error,
        severity: 'LOW'
      });
    }
  }

  /**
   * 指定されたスコープのすべてのキャッシュをクリアします
   * @param {string} scope キャッシュスコープ ('user', 'script', 'document')
   */
  static clear(scope = 'script') {
    try {
      // メモリキャッシュをクリア
      const memPrefix = `mem_${scope}_`;
      for (const key in this.memoryCache) {
        if (key.startsWith(memPrefix)) {
          delete this.memoryCache[key];
        }
      }

      // PropertiesServiceをクリア
      const props = this.getPropertiesByScope(scope);
      const allKeys = props.getKeys();
      const propPrefix = 'cache_';

      const keysToDelete = allKeys.filter(k => k.startsWith(propPrefix));
      for (const key of keysToDelete) {
        props.deleteProperty(key);
      }

      // CacheServiceは直接クリアできないが、スクリプトキャッシュの場合は可能
      if (scope === 'script') {
        CacheService.getScriptCache().removeAll([]);
      }
    } catch (error) {
      ErrorHandler.handleError({
        source: 'CacheManager.clear',
        error: error,
        severity: 'LOW'
      });
    }
  }

  /**
   * スコープに基づいてCacheServiceを取得します
   * @param {string} scope キャッシュスコープ
   * @return {Cache} Cache オブジェクト
   */
  static getCacheByScope(scope) {
    switch (scope) {
      case 'user':
        return CacheService.getUserCache();
      case 'document':
        return CacheService.getDocumentCache();
      case 'script':
      default:
        return CacheService.getScriptCache();
    }
  }

  /**
   * スコープに基づいてPropertiesServiceを取得します
   * @param {string} scope プロパティスコープ
   * @return {Properties} Properties オブジェクト
   */
  static getPropertiesByScope(scope) {
    switch (scope) {
      case 'user':
        return PropertiesService.getUserProperties();
      case 'document':
        return PropertiesService.getDocumentProperties();
      case 'script':
      default:
        return PropertiesService.getScriptProperties();
    }
  }
}

// メモリキャッシュの初期化
CacheManager.memoryCache = {};
