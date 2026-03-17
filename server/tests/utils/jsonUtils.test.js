import { describe, it, expect, vi } from 'vitest';
import { safeStringify, cleanObject, createSafeResponse } from '../../utils/jsonUtils.js';

describe('safeStringify', () => {
  it('stringifies a simple object', () => {
    const result = safeStringify({ a: 1, b: 'hello' });
    const parsed = JSON.parse(result);
    expect(parsed.a).toBe(1);
    expect(parsed.b).toBe('hello');
  });

  it('handles circular references', () => {
    const obj = { name: 'root' };
    obj.self = obj;
    const result = safeStringify(obj);
    const parsed = JSON.parse(result);
    expect(parsed.name).toBe('root');
    expect(parsed.self).toBe('[Circular Reference]');
  });

  it('converts undefined values to null', () => {
    const result = safeStringify({ a: undefined, b: 'valid' });
    const parsed = JSON.parse(result);
    expect(parsed.a).toBeNull();
    expect(parsed.b).toBe('valid');
  });

  it('converts functions to [Function]', () => {
    const result = safeStringify({ fn: () => {} });
    const parsed = JSON.parse(result);
    expect(parsed.fn).toBe('[Function]');
  });

  it('converts symbols to string representation', () => {
    const result = safeStringify({ sym: Symbol('test') });
    const parsed = JSON.parse(result);
    expect(parsed.sym).toBe('Symbol(test)');
  });

  it('converts BigInt to string', () => {
    const result = safeStringify({ big: BigInt(9007199254740991) });
    const parsed = JSON.parse(result);
    expect(parsed.big).toBe('9007199254740991');
  });

  it('respects custom indentation', () => {
    const result = safeStringify({ a: 1 }, 4);
    expect(result).toContain('    "a"');
  });

  it('handles null input', () => {
    const result = safeStringify(null);
    expect(result).toBe('null');
  });

  it('handles arrays', () => {
    const result = safeStringify([1, 2, 3]);
    expect(JSON.parse(result)).toEqual([1, 2, 3]);
  });
});

describe('cleanObject', () => {
  it('cleans a simple object', () => {
    const result = cleanObject({ name: 'test', value: 42 });
    expect(result).toEqual({ name: 'test', value: 42 });
  });

  it('removes underscore-prefixed properties', () => {
    const result = cleanObject({ name: 'test', _internal: 'hidden' });
    expect(result).toEqual({ name: 'test' });
    expect(result).not.toHaveProperty('_internal');
  });

  it('removes constructor and __proto__ properties', () => {
    const obj = { name: 'test', constructor: 'bad', __proto__: 'evil' };
    const result = cleanObject(obj);
    expect(result).not.toHaveProperty('constructor');
  });

  it('handles circular references', () => {
    const obj = { name: 'root' };
    obj.self = obj;
    const result = cleanObject(obj);
    expect(result.self).toBe('[Circular Reference]');
  });

  it('handles deeply nested objects up to maxDepth', () => {
    let obj = { value: 'leaf' };
    for (let i = 0; i < 15; i++) {
      obj = { child: obj };
    }
    const result = cleanObject(obj, 5);
    // Should have truncated deep nesting
    let current = result;
    let depth = 0;
    while (current && typeof current === 'object' && current.child) {
      current = current.child;
      depth++;
    }
    // At some point it should be '[Max depth exceeded]'
    expect(depth).toBeLessThanOrEqual(6);
  });

  it('handles null and undefined values', () => {
    expect(cleanObject(null)).toBeNull();
    expect(cleanObject(undefined)).toBeNull();
  });

  it('handles arrays', () => {
    const result = cleanObject([1, { name: 'test', _private: 'hidden' }, 3]);
    expect(result).toEqual([1, { name: 'test' }, 3]);
  });

  it('returns primitive values as-is', () => {
    expect(cleanObject(42)).toBe(42);
    expect(cleanObject('string')).toBe('string');
    expect(cleanObject(true)).toBe(true);
  });
});

describe('createSafeResponse', () => {
  it('creates MCP-compliant response with cleaned and stringified data', () => {
    const result = createSafeResponse({ message: 'success', count: 5 });
    expect(result.content).toHaveLength(1);
    expect(result.content[0].type).toBe('text');
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.message).toBe('success');
    expect(parsed.count).toBe(5);
  });

  it('handles circular references in response data', () => {
    const data = { name: 'test' };
    data.self = data;
    const result = createSafeResponse(data);
    expect(result.content[0].type).toBe('text');
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.self).toBe('[Circular Reference]');
  });

  it('skips cleaning when clean option is false', () => {
    const result = createSafeResponse({ _private: 'visible' }, { clean: false });
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed._private).toBe('visible');
  });

  it('strips underscore properties by default', () => {
    const result = createSafeResponse({ name: 'test', _hidden: 'secret' });
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed).not.toHaveProperty('_hidden');
  });
});
