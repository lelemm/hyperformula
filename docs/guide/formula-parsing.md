# Formula parsing

Use `tokenizeFormula()` and `parseFormula()` when your application needs formula tokens or AST data, for example to power code editor decorations, syntax highlighting, autocomplete, or other formula-aware tooling.

Both methods are instance methods, so their output follows the instance configuration: localized function names, separators, registered functions, and sheet names.

## Tokenizing a formula

`tokenizeFormula()` returns plain token objects and lexing errors through a stable HyperFormula API contract, even if the underlying lexer implementation changes.

```js
const hfInstance = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });

const tokenization = hfInstance.tokenizeFormula('=SUM(A1, 2)');

console.log(tokenization.tokens.map((token) => token.image));
// ['=', 'SUM(', 'A1', ',', ' ', '2', ')']
```

## Parsing a formula

`parseFormula()` returns a stable public AST, parsing errors, dependencies, and flags for volatile or structural-change functions.

```js
const hfInstance = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });

const parsing = hfInstance.parseFormula('=SUM(A1, 2)');

console.log(parsing.ast.type);
// 'FUNCTION_CALL'

console.log(parsing.dependencies);
// [{ type: 'CELL', address: { sheet: 0, col: 0, row: 0 } }]
```

Pass a `SimpleCellAddress` as the second argument when relative references should be parsed in a specific sheet context.

```js
const parsing = hfInstance.parseFormula('=A1', { sheet: 1, col: 3, row: 3 });
```

Unlike the internal parser used while building dependency graphs, `parseFormula()` doesn't create placeholder sheets when a formula references an unknown sheet.

## Invalid input

Both methods require a formula string that starts with `=`. If the input is not a formula, they throw `NotAFormulaError`.
