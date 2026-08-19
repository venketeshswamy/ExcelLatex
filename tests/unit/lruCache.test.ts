import { describe, it, expect } from 'vitest';
import { LRUCache, createRenderCacheKey } from '../../src/core/lruCache';

describe('lruCache.ts - LRUCache and Cache Key Generation', () => {
  it('should store and retrieve values correctly', () => {
    const cache = new LRUCache<string, number>({ maxSize: 3 });
    cache.set('a', 1);
    cache.set('b', 2);

    expect(cache.get('a')).toBe(1);
    expect(cache.get('b')).toBe(2);
    expect(cache.get('c')).toBeUndefined();
  });

  it('should evict the least recently used item when reaching capacity', () => {
    const cache = new LRUCache<string, string>({ maxSize: 3 });
    cache.set('a', 'alpha');
    cache.set('b', 'beta');
    cache.set('c', 'gamma');

    // Access 'a' to make it recently used (order now: b, c, a)
    expect(cache.get('a')).toBe('alpha');

    // Adding 'd' should evict 'b'
    cache.set('d', 'delta');

    expect(cache.has('b')).toBe(false);
    expect(cache.get('b')).toBeUndefined();
    expect(cache.get('a')).toBe('alpha');
    expect(cache.get('c')).toBe('gamma');
    expect(cache.get('d')).toBe('delta');
    expect(cache.size()).toBe(3);
  });

  it('should support item deletion and clearing', () => {
    const cache = new LRUCache<string, number>({ maxSize: 5 });
    cache.set('x', 10);
    cache.set('y', 20);

    expect(cache.delete('x')).toBe(true);
    expect(cache.has('x')).toBe(false);
    expect(cache.size()).toBe(1);

    cache.clear();
    expect(cache.size()).toBe(0);
    expect(cache.has('y')).toBe(false);
  });

  it('should track cache hits, misses, and hit ratio accurately', () => {
    const cache = new LRUCache<string, number>({ maxSize: 2 });
    cache.set('k1', 100);

    // Hit
    cache.get('k1');
    // Miss
    cache.get('k2');
    // Eviction
    cache.set('k2', 200);
    cache.set('k3', 300); // evicts k1

    const stats = cache.getStats();
    expect(stats.hits).toBe(1);
    expect(stats.misses).toBe(1);
    expect(stats.evictions).toBe(1);
    expect(stats.totalRequests).toBe(2);
    expect(stats.hitRatio).toBe(0.5);
  });

  it('should handle TTL expiration', async () => {
    const cache = new LRUCache<string, string>({ maxSize: 5, defaultTtlMs: 50 });
    cache.set('short', 'lived');

    expect(cache.get('short')).toBe('lived');

    // Wait for expiration
    await new Promise((resolve) => setTimeout(resolve, 70));

    expect(cache.get('short')).toBeUndefined();
    expect(cache.has('short')).toBe(false);
  });

  it('should create deterministic cache keys', () => {
    const key1 = createRenderCacheKey('\\frac{a}{b}', {
      background: 'transparent',
      color: '#000000',
      fontSize: 16,
      displayMode: true,
      scale: 3
    });
    const key2 = createRenderCacheKey(' \\frac{a}{b} ', {
      background: 'TRANSPARENT',
      color: '#000000',
      fontSize: 16,
      displayMode: true,
      scale: 3
    });
    expect(key1).toBe(key2);
  });
});
