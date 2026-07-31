/**
 * Generates a fast, low-overhead 32-bit deterministic numeric hash sequence from a string payload.
 * Leverages a customized variant of the Jenkins / DJB2 string folding rotation algorithm to keep
 * CPU execution bounds tight inside deep iterative styling loops.
 *
 * @param val Input string payload configuration to fingerprint
 */
export function fastHashString(val: string): string {
  let hash = 5381;
  let i = val.length;

  while (i--) {
    // Perform bitwise bit-shift wrapping operations sequentially
    hash = (hash * 33) ^ val.charCodeAt(i);
  }

  // Convert the unsigned 32-bit integer result to a base-36 uppercase layout token string representation
  return (hash >>> 0).toString(36).toUpperCase();
}

/**
 * Deep-inspects an object structure cleanly to emit a reliable deterministic fingerprint lookup signature string.
 */
export function generateObjectSignature(obj: Record<string, any> | null | undefined): string {
  if (!obj) return "EMPTY_SIG";

  // JSON stringify provides stable serialization tracking since configuration contracts use standardized typing shapes
  return fastHashString(JSON.stringify(obj));
}
