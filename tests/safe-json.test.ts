import assert from "node:assert/strict";
import { test } from "node:test";

import { SafeInputError, parseJson } from "../src/safe-json.js";

/** The JSON reader's own properties, tested directly since `parseSnapshot`'s tighter shape rules would hide a reader that stopped dropping a reserved name or counted characters where it says bytes. */

/** This project's own bounds are `parseSnapshot`'s business; these are the reader's. */
const BOUNDS = { maxBytes: 4096, maxDepth: 8, maxKeys: 100 };
/** Distinctive, because the reader promises this opens every refusal it writes. */
const WHAT = "The stored copy";
const OPTIONS = { ...BOUNDS, what: WHAT };

const read = (text: string, options: Partial<typeof OPTIONS> = {}) =>
  parseJson(text, { ...OPTIONS, ...options }) as Record<string, unknown>;

test("the reader drops the names the object model reserves, and nothing else", () => {
  // JSON string escapes decode before the reviver sees the key, so this payload arrives as `__proto__` and is dropped — which is why the defence lives in the reviver, not a scan of the raw text.
  const escaped =
    '{"\\u005f\\u005fproto\\u005f\\u005f":{"pwned":true},"keep":1}';
  assert.equal(
    escaped.includes("__proto__"),
    false,
    "the payload must not spell the name it smuggles, or it proves nothing",
  );
  const decoded = read(escaped);
  assert.deepEqual(Object.keys(decoded), ["keep"]);
  assert.equal(Object.getPrototypeOf(decoded), null);
  assert.equal(
    ({} as Record<string, unknown>).pwned,
    undefined,
    "every object in this process now has a pwned property",
  );

  // The plain spelling, and the two other reserved names, at the root and nested — the reviver visits every level.
  const plain = read(
    '{"__proto__":{"pwned":true},"constructor":1,"prototype":2,"nested":{"__proto__":{},"constructor":3,"kept":4}}',
  );
  assert.deepEqual(Object.keys(plain), ["nested"]);
  assert.deepEqual(Object.keys(plain.nested as object), ["kept"]);

  // AND NOTHING ELSE: the reader promises a SET of three reserved names, so a fourth name — even a lookalike — is data, not a threat.
  const lookalikes = read(
    '{"__note__":1,"__proto":2,"proto__":3,"_constructor":4,"Constructor":5,"PROTOTYPE":6}',
  );
  assert.deepEqual(Object.keys(lookalikes), [
    "__note__",
    "__proto",
    "proto__",
    "_constructor",
    "Constructor",
    "PROTOTYPE",
  ]);
});

test("the cap the reader enforces is on bytes, not on characters", () => {
  // `maxBytes` bounds what a store or file will hold; measuring `text.length` instead would admit up to three times the intended bytes.
  const wide = `{"a":"${"é".repeat(6)}"}`;
  assert.equal(wide.length, 14);
  assert.equal(new TextEncoder().encode(wide).length, 20);
  assert.throws(
    () => read(wide, { maxBytes: 15 }),
    {
      name: "SafeInputError",
      // Bytes, and the exact count: rounded to megabytes, a file one byte over its cap used to be refused for being exactly the size it is allowed.
      message: new RegExp(`^${WHAT} is 20 bytes, over the 15-byte limit\\.$`),
    },
    "20 bytes passed a 15-byte cap, so the reader is counting characters",
  );

  // The control: the same 14 characters, one byte each, are inside the cap, so the case above is about bytes, not string length.
  const narrow = `{"a":"${"e".repeat(6)}"}`;
  assert.equal(narrow.length, wide.length);
  assert.equal(read(narrow, { maxBytes: 15 }).a, "eeeeee");
});

test("a top-level scalar is refused rather than handed back", () => {
  // `null` is worth testing because `typeof null` is "object", so the obvious caller guard lets it through.
  for (const scalar of ["42", '"text"', "null", "true"]) {
    assert.throws(
      () => read(scalar),
      {
        name: "SafeInputError",
        message: `${WHAT} is not a JSON object or array.`,
      },
      `${scalar} was accepted as a document`,
    );
  }

  // Both shapes ARE documents; objects come back on a null prototype, but an array comes back a plain array (a null-prototype array is not one).
  assert.equal(Object.getPrototypeOf(read("{}")), null);
  const array = parseJson("[1,2]", OPTIONS);
  assert.ok(Array.isArray(array));
  assert.deepEqual(array, [1, 2]);
});

test("every refusal the reader writes is a sentence naming what was read", () => {
  // `what` is the caller's word for the thing the person chose, and it opens every refusal so it isn't read as a complaint about a different file.
  for (const [payload, bounds] of [
    ["", {}],
    ["{ not json", {}],
    ["[[[[[[[[[[]]]]]]]]]]", {}],
    ['{"a":1,"b":2,"c":3}', { maxKeys: 2 }],
  ] as const) {
    assert.throws(
      () => read(payload, bounds),
      (error: unknown) => {
        assert.ok(
          error instanceof SafeInputError,
          `${payload}: not a SafeInputError`,
        );
        assert.ok(
          error.message.startsWith(WHAT),
          `refusal does not name what was read: ${error.message}`,
        );
        return true;
      },
    );
  }
});
