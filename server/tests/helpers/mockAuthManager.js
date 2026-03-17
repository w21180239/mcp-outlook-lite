import { vi } from 'vitest';

export function createMockAuthManager(overrides = {}) {
  const mockGraphClient = {
    api: vi.fn().mockReturnThis(),
    get: vi.fn().mockResolvedValue({}),
    post: vi.fn().mockResolvedValue({}),
    patch: vi.fn().mockResolvedValue({}),
    delete: vi.fn().mockResolvedValue({}),
    select: vi.fn().mockReturnThis(),
    filter: vi.fn().mockReturnThis(),
    top: vi.fn().mockReturnThis(),
    skip: vi.fn().mockReturnThis(),
    orderby: vi.fn().mockReturnThis(),
    expand: vi.fn().mockReturnThis(),
    count: vi.fn().mockReturnThis(),
    header: vi.fn().mockReturnThis(),
    headers: vi.fn().mockReturnThis(),
    ...overrides,
  };

  return {
    ensureAuthenticated: vi.fn().mockResolvedValue(undefined),
    getGraphClient: vi.fn().mockReturnValue(mockGraphClient),
    getGraphApiClient: vi.fn().mockReturnValue({
      makeRequest: vi.fn().mockResolvedValue({}),
      postWithRetry: vi.fn().mockResolvedValue({}),
      deleteWithRetry: vi.fn().mockResolvedValue({}),
    }),
    graphClient: mockGraphClient,
    graphApiClient: {
      makeRequest: vi.fn().mockResolvedValue({}),
    },
    isAuthenticated: true,
  };
}
