import { vi } from 'vitest';

/**
 * Creates an object where every listed method is a `vi.fn()` spy.
 * Drop-in replacement for Jasmine's `jasmine.createSpyObj(name, methods)`.
 */
export function createSpyObj<T extends string>(
    _name: string,
    methods: T[]
): Record<T, ReturnType<typeof vi.fn>> {
    return Object.fromEntries(methods.map(m => [m, vi.fn()])) as Record<
        T,
        ReturnType<typeof vi.fn>
    >;
}
