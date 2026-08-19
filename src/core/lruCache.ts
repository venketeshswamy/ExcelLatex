/**
 * High-Performance In-Memory LRU (Least Recently Used) Cache.
 * Provides O(1) time complexity for get, set, has, and delete operations.
 */

export interface LRUCacheOptions {
  maxSize?: number;
  defaultTtlMs?: number; // Optional TTL in milliseconds
}

export interface CacheStats {
  size: number;
  maxSize: number;
  hits: number;
  misses: number;
  evictions: number;
  totalRequests: number;
  hitRatio: number;
}

interface CacheNode<K, V> {
  key: K;
  value: V;
  expiresAt: number | null;
  prev: CacheNode<K, V> | null;
  next: CacheNode<K, V> | null;
}

export class LRUCache<K, V> {
  private readonly maxSize: number;
  private readonly defaultTtlMs: number | null;
  private lookup: Map<K, CacheNode<K, V>> = new Map();
  private head: CacheNode<K, V> | null = null;
  private tail: CacheNode<K, V> | null = null;

  private hits = 0;
  private misses = 0;
  private evictions = 0;

  constructor(options: LRUCacheOptions = {}) {
    this.maxSize = Math.max(1, options.maxSize ?? 500);
    this.defaultTtlMs = options.defaultTtlMs && options.defaultTtlMs > 0 ? options.defaultTtlMs : null;
  }

  public get(key: K): V | undefined {
    const node = this.lookup.get(key);
    if (!node) {
      this.misses++;
      return undefined;
    }

    if (node.expiresAt !== null && Date.now() > node.expiresAt) {
      this.delete(key);
      this.misses++;
      return undefined;
    }

    this.hits++;
    this.moveToHead(node);
    return node.value;
  }

  public set(key: K, value: V, ttlMs?: number): void {
    const effectiveTtl = ttlMs ?? this.defaultTtlMs;
    const expiresAt = effectiveTtl ? Date.now() + effectiveTtl : null;

    let node = this.lookup.get(key);
    if (node) {
      node.value = value;
      node.expiresAt = expiresAt;
      this.moveToHead(node);
      return;
    }

    if (this.lookup.size >= this.maxSize) {
      this.evictTail();
    }

    node = {
      key,
      value,
      expiresAt,
      prev: null,
      next: null
    };

    this.lookup.set(key, node);
    this.addToHead(node);
  }

  public has(key: K): boolean {
    const node = this.lookup.get(key);
    if (!node) return false;
    if (node.expiresAt !== null && Date.now() > node.expiresAt) {
      this.delete(key);
      return false;
    }
    return true;
  }

  public delete(key: K): boolean {
    const node = this.lookup.get(key);
    if (!node) return false;
    this.removeNode(node);
    this.lookup.delete(key);
    return true;
  }

  public clear(): void {
    this.lookup.clear();
    this.head = null;
    this.tail = null;
    this.hits = 0;
    this.misses = 0;
    this.evictions = 0;
  }

  public size(): number {
    return this.lookup.size;
  }

  public getStats(): CacheStats {
    const totalRequests = this.hits + this.misses;
    const hitRatio = totalRequests > 0 ? this.hits / totalRequests : 0;
    return {
      size: this.lookup.size,
      maxSize: this.maxSize,
      hits: this.hits,
      misses: this.misses,
      evictions: this.evictions,
      totalRequests,
      hitRatio
    };
  }

  private addToHead(node: CacheNode<K, V>): void {
    node.next = this.head;
    node.prev = null;
    if (this.head) {
      this.head.prev = node;
    }
    this.head = node;
    if (!this.tail) {
      this.tail = node;
    }
  }

  private removeNode(node: CacheNode<K, V>): void {
    if (node.prev) {
      node.prev.next = node.next;
    } else {
      this.head = node.next;
    }

    if (node.next) {
      node.next.prev = node.prev;
    } else {
      this.tail = node.prev;
    }

    node.prev = null;
    node.next = null;
  }

  private moveToHead(node: CacheNode<K, V>): void {
    if (this.head === node) return;
    this.removeNode(node);
    this.addToHead(node);
  }

  private evictTail(): void {
    if (!this.tail) return;
    const oldTail = this.tail;
    this.removeNode(oldTail);
    this.lookup.delete(oldTail.key);
    this.evictions++;
  }
}

/**
 * Creates a deterministic cache key for LaTeX rendering options.
 */
export function createRenderCacheKey(
  latex: string,
  options?: {
    background?: string | number;
    color?: string;
    fontSize?: number;
    displayMode?: boolean;
    scale?: number;
  }
): string {
  const bg = String(options?.background ?? 'transparent').trim().toLowerCase();
  const color = (options?.color ?? '#000000').trim().toLowerCase();
  const size = options?.fontSize ?? 16;
  const display = options?.displayMode ?? true;
  const scale = options?.scale ?? 3;
  return `${latex.trim()}||${bg}||${color}||${size}||${display}||${scale}`;
}
