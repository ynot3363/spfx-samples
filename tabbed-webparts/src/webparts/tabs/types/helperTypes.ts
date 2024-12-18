/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * Creates a union of string literal types representing the paths to properties within an object.
 * It recursively traverses the object's properties until it encounters an array or an undefined value.
 * For arrays, it stops the recursion and includes the key for the array itself in the union.
 * For objects, it includes both the key and the paths to deeper properties.
 * It excludes paths leading to undefined values.
 * @template ObjectType - The object type to extract nested keys from.
 * @returns A union of string literal types representing the paths to properties
 */

export type NestedKeyOf<ObjectType> = {
  //Iterate over each key of the object that is a string or a number
  [Key in keyof ObjectType &
    (string | number)]: ObjectType[Key] extends Array<any> //Check if the propert at this key is an array
    ? `${Key}` // If it's an array, include just the key, stop the recursion.
    : ObjectType[Key] extends object | undefined // Check if the property is an object or potentially undefined
    ? ObjectType[Key] extends undefined
      ? never // If it's undefined, exclude this path
      : `${Key}` | `${Key}.${NestedKeyOf<NonNullable<ObjectType[Key]>>}` //If it is an object, include the key and recurse into it for deeper paths
    : `${Key}`; // For all other types, include the Key
}[keyof ObjectType & (string | number)]; // Compile the resulting types

/**
 * Recursively resolves the type of a value at a given path within an object.
 * It handles direct key access, nested paths and potentially undefined values.
 * For nest paths, it infers the property type at each level and proceeds until the end of the path.
 * If it encounters an undefined value, it stope and does not return a type.
 * @template T - The object type to extract the value type from.
 * @template K - The string representing the path to the value within the object.
 * @returns The type of the value located at the specified path within the object or never if the path is invalid or leads to an undefined value.
 */
export type NestedValueType<T, K extends string> = K extends keyof T
  ? T[K] // Direct key access: If K is a direct key of T, return its type
  : K extends `${infer P}.${infer R}` // Nested paths: If K represents a nested path, split it into parts.
  ? P extends keyof T
    ? T[P] extends infer TP // Infer the property type of the current segement
      ? TP extends undefined // Check if the current segement might be undefined
        ? never // If the current segment is strictly undefined, return never
        : NestedValueType<NonNullable<TP>, R> // If it is not undefined, continue with the next segment
      : never // IF the propert type could not be inferred, return never
    : never // If the current segment is not a key of T, return never
  : never; // If K is not a key of T or a valid nested path, return never
