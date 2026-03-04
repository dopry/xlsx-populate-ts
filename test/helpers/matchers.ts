import { expect } from 'vitest';
import jsondiffpatch from 'jsondiffpatch';

interface CustomMatchers<R = unknown> {
    toEqualJson(expected: unknown): R;
    toEqualUInt8Array(expected: Uint8Array): R;
}

declare module 'vitest' {
    interface Assertion<T = any> extends CustomMatchers<T> {}
    interface AsymmetricMatchersContaining extends CustomMatchers {}
}

expect.extend({
    toEqualJson(received: unknown, expected: unknown) {
        const cleanReceived = JSON.parse(JSON.stringify(received));
        const cleanExpected = JSON.parse(JSON.stringify(expected));
        // @ts-ignore – vitest provides this.equals() on the context
        const pass = this.equals(cleanReceived, cleanExpected);
        return {
            pass,
            message: () =>
                pass
                    ? "Expected objects NOT to equal as JSON"
                    : "JSON objects don't match:\n" +
                      jsondiffpatch.formatters.console.format(
                          jsondiffpatch.diff(cleanExpected, cleanReceived)
                      ),
        };
    },

    toEqualUInt8Array(received: Uint8Array, expected: Uint8Array) {
        let pass = received.byteLength === expected.byteLength;
        if (pass) {
            for (let i = 0; i < received.byteLength; i++) {
                if (received[i] !== expected[i]) {
                    pass = false;
                    break;
                }
            }
        }
        return {
            pass,
            message: () =>
                pass ? "Expected UInt8Arrays NOT to match" : "UInt8Arrays do not match",
        };
    },
});
