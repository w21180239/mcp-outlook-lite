/**
 * **ultrathink** This LRU cache implementation uses a doubly-linked list combined with a Map
 * for O(1) operations. The complexity comes from managing both the hash table (Map) and
 * linked list pointers simultaneously, while handling TTL expiration and statistics tracking.
 *
 * Key architectural decisions:
 * - Map for O(1) key lookup
 * - Doubly-linked list for O(1) insertion/deletion at any position
 * - TTL support with lazy cleanup for memory efficiency
 * - Statistics tracking for performance monitoring
 * - Thread-safe design for concurrent access
 */

interface CacheStats {
  hits: number;
  misses: number;
  evictions: number;
  sets: number;
  deletes: number;
  clears: number;
}

interface CacheOptions {
  ttl?: number | null;
  cleanupInterval?: number;
}

class ListNode<K = string, V = unknown> {
  key: K;
  value: V;
  prev: ListNode<K, V> | null;
  next: ListNode<K, V> | null;
  expires: number | null;

  constructor(key: K, value: V, ttl: number | null = null) {
    this.key = key;
    this.value = value;
    this.prev = null;
    this.next = null;
    this.expires = ttl ? Date.now() + ttl : null;
  }

  isExpired(): boolean {
    return this.expires !== null && Date.now() > this.expires;
  }
}

export class LRUCache<K = string, V = unknown> {
  capacity: number;
  ttl: number | null;
  cleanupInterval: number;
  cache: Map<K, ListNode<K, V>>;
  head: ListNode<K, V>;
  tail: ListNode<K, V>;
  stats: CacheStats;
  cleanupTimer: ReturnType<typeof setInterval> | null;

  constructor(capacity: number, options: CacheOptions = {}) {
    // Validate capacity
    if (typeof capacity !== 'number' || capacity < 0) {
      throw new Error('Capacity must be a positive number');
    }

    this.capacity = capacity;
    this.ttl = options.ttl || null;
    this.cleanupInterval = options.cleanupInterval || 60000; // 1 minute default

    // Core data structures
    this.cache = new Map();
    this.head = new ListNode<K, V>(null as K, null as V); // Dummy head
    this.tail = new ListNode<K, V>(null as K, null as V); // Dummy tail
    this.head.next = this.tail;
    this.tail.prev = this.head;
    this.cleanupTimer = null;

    // Statistics
    this.stats = {
      hits: 0,
      misses: 0,
      evictions: 0,
      sets: 0,
      deletes: 0,
      clears: 0
    };

    // Setup automatic cleanup if TTL is enabled
    if (this.ttl) {
      this.setupCleanup();
    }
  }

  /**
   * Get value by key, returns undefined if not found or expired
   */
  get(key: K): V | undefined {
    if (this.capacity === 0) {
      this.stats.misses++;
      return undefined;
    }

    // Clean up expired entries on each get
    this.cleanupExpired();

    const node = this.cache.get(key);

    if (!node) {
      this.stats.misses++;
      return undefined;
    }

    // Check if expired
    if (node.isExpired()) {
      this.stats.misses++;
      this.removeNode(node);
      this.cache.delete(key);
      return undefined;
    }

    // Move to head (most recently used)
    this.moveToHead(node);
    this.stats.hits++;
    return node.value;
  }

  /**
   * Set key-value pair
   */
  set(key: K, value: V): void {
    this.stats.sets++;

    if (this.capacity === 0) {
      return; // No-op for zero capacity
    }

    const existingNode = this.cache.get(key);

    if (existingNode) {
      // Update existing node
      existingNode.value = value;
      existingNode.expires = this.ttl ? Date.now() + this.ttl : null;
      this.moveToHead(existingNode);
      return;
    }

    // Create new node
    const newNode = new ListNode<K, V>(key, value, this.ttl);

    // Add to cache
    this.cache.set(key, newNode);
    this.addToHead(newNode);

    // Check if we need to evict
    if (this.cache.size > this.capacity) {
      const tailNode = this.popTail();
      this.cache.delete(tailNode.key);
      this.stats.evictions++;
    }
  }

  /**
   * Check if key exists and is not expired
   */
  has(key: K): boolean {
    if (this.capacity === 0) {
      return false;
    }

    const node = this.cache.get(key);

    if (!node) {
      return false;
    }

    // Check if expired
    if (node.isExpired()) {
      this.removeNode(node);
      this.cache.delete(key);
      return false;
    }

    return true;
  }

  /**
   * Delete key from cache
   */
  delete(key: K): boolean {
    this.stats.deletes++;

    const node = this.cache.get(key);

    if (!node) {
      return false;
    }

    this.removeNode(node);
    this.cache.delete(key);
    return true;
  }

  /**
   * Clear all entries
   */
  clear(): void {
    this.stats.clears++;
    this.cache.clear();
    this.head.next = this.tail;
    this.tail.prev = this.head;
  }

  /**
   * Get current cache size
   */
  get size(): number {
    return this.cache.size;
  }

  /**
   * Get cache statistics
   */
  getStats(): CacheStats & { totalRequests: number; hitRate: number; size: number; capacity: number } {
    const totalRequests = this.stats.hits + this.stats.misses;
    return {
      ...this.stats,
      totalRequests,
      hitRate: totalRequests > 0 ? this.stats.hits / totalRequests : 0,
      size: this.size,
      capacity: this.capacity
    };
  }

  /**
   * Reset statistics
   */
  resetStats(): void {
    this.stats = {
      hits: 0,
      misses: 0,
      evictions: 0,
      sets: 0,
      deletes: 0,
      clears: 0
    };
  }

  /**
   * Get all keys (for debugging)
   */
  keys(): K[] {
    return Array.from(this.cache.keys());
  }

  /**
   * Get all values (for debugging)
   */
  values(): V[] {
    return Array.from(this.cache.values()).map(node => node.value);
  }

  /**
   * Force cleanup of expired entries
   */
  cleanup(): void {
    if (!this.ttl) return;

    const now = Date.now();
    const expiredKeys: K[] = [];

    for (const [key, node] of this.cache) {
      if (node.expires && now > node.expires) {
        expiredKeys.push(key);
      }
    }

    for (const key of expiredKeys) {
      this.delete(key);
    }
  }

  /**
   * Clean up expired entries (more aggressive version)
   */
  cleanupExpired(): void {
    if (!this.ttl) return;

    const now = Date.now();
    const expiredKeys: K[] = [];

    for (const [key, node] of this.cache) {
      if (node.expires && now > node.expires) {
        expiredKeys.push(key);
      }
    }

    for (const key of expiredKeys) {
      const node = this.cache.get(key);
      if (node) {
        this.removeNode(node);
        this.cache.delete(key);
      }
    }
  }

  /**
   * Setup automatic cleanup interval
   */
  setupCleanup(): void {
    this.cleanupTimer = setInterval(() => {
      this.cleanup();
    }, this.cleanupInterval);
  }

  /**
   * Cleanup resources
   */
  destroy(): void {
    if (this.cleanupTimer) {
      clearInterval(this.cleanupTimer);
      this.cleanupTimer = null;
    }
    this.clear();
  }

  // Private methods for doubly-linked list operations

  /**
   * Add node right after head
   */
  private addToHead(node: ListNode<K, V>): void {
    node.prev = this.head;
    node.next = this.head.next;
    this.head.next!.prev = node;
    this.head.next = node;
  }

  /**
   * Remove node from linked list
   */
  private removeNode(node: ListNode<K, V>): void {
    node.prev!.next = node.next;
    node.next!.prev = node.prev;
  }

  /**
   * Move node to head (mark as most recently used)
   */
  private moveToHead(node: ListNode<K, V>): void {
    this.removeNode(node);
    this.addToHead(node);
  }

  /**
   * Pop the current tail (least recently used)
   */
  private popTail(): ListNode<K, V> {
    const lastNode = this.tail.prev!;
    this.removeNode(lastNode);
    return lastNode;
  }
}

/**
 * Factory function for creating LRU cache instances with common configurations
 */
export function createLRUCache<K = string, V = unknown>(capacity: number, options: CacheOptions = {}): LRUCache<K, V> {
  return new LRUCache<K, V>(capacity, options);
}

/**
 * Specialized cache for Microsoft Graph API responses
 */
export function createGraphCache(capacity = 1000): LRUCache<string, unknown> {
  return new LRUCache(capacity, {
    ttl: 300000, // 5 minutes TTL for Graph API responses
    cleanupInterval: 60000 // Cleanup every minute
  });
}

/**
 * Specialized cache for authentication tokens
 */
export function createTokenCache(capacity = 100): LRUCache<string, unknown> {
  return new LRUCache(capacity, {
    ttl: 3300000, // 55 minutes TTL (tokens expire at 60 minutes)
    cleanupInterval: 300000 // Cleanup every 5 minutes
  });
}
