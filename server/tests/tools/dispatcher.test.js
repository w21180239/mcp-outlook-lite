import { describe, it, expect } from 'vitest';
import { getToolHandler, getRegisteredToolNames } from '../../tools/dispatcher.js';
import { allToolSchemas } from '../../schemas/toolSchemas.js';

describe('dispatcher', () => {
  describe('getToolHandler', () => {
    it('should return a function for every registered tool name', () => {
      const names = getRegisteredToolNames();
      expect(names.length).toBeGreaterThan(0);

      for (const name of names) {
        const handler = getToolHandler(name);
        expect(handler, `Handler for ${name} should be a function`).toBeTypeOf('function');
      }
    });

    it('should return null for unknown tool name', () => {
      expect(getToolHandler('nonexistent_tool')).toBeNull();
    });

    it('should return null for empty string', () => {
      expect(getToolHandler('')).toBeNull();
    });
  });

  describe('schema-registry alignment', () => {
    it('should have a handler for every schema tool name', () => {
      const schemaNames = allToolSchemas.map(s => s.name);
      const registeredNames = getRegisteredToolNames();

      for (const schemaName of schemaNames) {
        expect(registeredNames, `Missing handler for schema: ${schemaName}`).toContain(schemaName);
      }
    });

    it('should not have handlers for tools without schemas', () => {
      const schemaNames = new Set(allToolSchemas.map(s => s.name));
      const registeredNames = getRegisteredToolNames();

      for (const name of registeredNames) {
        expect(schemaNames.has(name), `Handler ${name} has no matching schema`).toBe(true);
      }
    });
  });
});
