import { describe, it, expect } from 'vitest';
import { getSuccessPage, getErrorPage, getFailurePage } from '../../auth/templates.js';

describe('auth templates', () => {
  describe('getSuccessPage', () => {
    it('should return valid HTML with closing tag', () => {
      const html = getSuccessPage();
      expect(html).toContain('</html>');
    });

    it('should contain success messaging', () => {
      const html = getSuccessPage();
      expect(html).toContain('Authentication Successful');
    });

    it('should contain auto-close countdown script', () => {
      const html = getSuccessPage();
      expect(html).toContain('countdown');
      expect(html).toContain('window.close');
    });

    it('should contain icon', () => {
      const html = getSuccessPage();
      expect(html).toContain('class="icon"');
    });
  });

  describe('getErrorPage', () => {
    it('should return valid HTML with closing tag', () => {
      const html = getErrorPage();
      expect(html).toContain('</html>');
    });

    it('should contain security error messaging', () => {
      const html = getErrorPage();
      expect(html).toContain('Security Error');
    });

    it('should contain reconnect instructions', () => {
      const html = getErrorPage();
      expect(html).toContain('disconnect and reconnect');
    });
  });

  describe('getFailurePage', () => {
    it('should return valid HTML with closing tag', () => {
      const html = getFailurePage();
      expect(html).toContain('</html>');
    });

    it('should contain failure messaging', () => {
      const html = getFailurePage();
      expect(html).toContain('Authentication Failed');
    });

    it('should contain reconnect instructions', () => {
      const html = getFailurePage();
      expect(html).toContain('disconnect and reconnect');
    });
  });
});
