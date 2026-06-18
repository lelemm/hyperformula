/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {ILexingError, ILexingResult, IToken} from 'chevrotain'
import {SimpleCellAddress} from './Cell'
import {
  AddressDependency,
  Ast,
  AstNodeType,
  CellRangeDependency,
  ColumnRangeDependency,
  NamedExpressionDependency,
  RelativeDependency,
  RowRangeDependency,
} from './parser'
import {CellAddress} from './parser/CellAddress'
import {ColumnAddress} from './parser/ColumnAddress'
import {ParsingError, RangeSheetReferenceType} from './parser/Ast'
import {ParsingResult} from './parser/ParserWithCaching'
import {RowAddress} from './parser/RowAddress'

export interface FormulaTokenizationResult {
  tokens: FormulaToken[],
  errors: FormulaLexingError[],
}

export interface FormulaToken {
  type: string,
  image: string,
  startOffset: number,
  endOffset: number,
  startLine?: number,
  endLine?: number,
  startColumn?: number,
  endColumn?: number,
}

export interface FormulaLexingError {
  message: string,
  offset: number,
  line?: number,
  column?: number,
  length?: number,
}

export interface FormulaParsingResult {
  ast?: FormulaAstNode,
  errors: FormulaParsingError[],
  dependencies: FormulaDependency[],
  hasVolatileFunction: boolean,
  hasStructuralChangeFunction: boolean,
}

export interface FormulaParsingError {
  type: string,
  message: string,
}

export type FormulaAstNode =
  FormulaEmptyAstNode
  | FormulaNumberAstNode
  | FormulaStringAstNode
  | FormulaCellReferenceAstNode
  | FormulaCellRangeAstNode
  | FormulaColumnRangeAstNode
  | FormulaRowRangeAstNode
  | FormulaBinaryOpAstNode
  | FormulaUnaryOpAstNode
  | FormulaFunctionCallAstNode
  | FormulaNamedExpressionAstNode
  | FormulaParenthesisAstNode
  | FormulaErrorAstNode
  | FormulaErrorWithRawInputAstNode
  | FormulaArrayAstNode

export interface FormulaAstNodeBase {
  type: string,
  leadingWhitespace?: string,
}

export interface FormulaEmptyAstNode extends FormulaAstNodeBase {
  type: 'EMPTY',
}

export interface FormulaNumberAstNode extends FormulaAstNodeBase {
  type: 'NUMBER',
  value: number,
}

export interface FormulaStringAstNode extends FormulaAstNodeBase {
  type: 'STRING',
  value: string,
}

export interface FormulaCellReferenceAstNode extends FormulaAstNodeBase {
  type: 'CELL_REFERENCE',
  reference: FormulaCellReference,
}

export interface FormulaCellRangeAstNode extends FormulaAstNodeBase {
  type: 'CELL_RANGE',
  start: FormulaCellReference,
  end: FormulaCellReference,
  sheetReferenceType: FormulaRangeSheetReferenceType,
}

export interface FormulaColumnRangeAstNode extends FormulaAstNodeBase {
  type: 'COLUMN_RANGE',
  start: FormulaColumnReference,
  end: FormulaColumnReference,
  sheetReferenceType: FormulaRangeSheetReferenceType,
}

export interface FormulaRowRangeAstNode extends FormulaAstNodeBase {
  type: 'ROW_RANGE',
  start: FormulaRowReference,
  end: FormulaRowReference,
  sheetReferenceType: FormulaRangeSheetReferenceType,
}

export type FormulaBinaryOpType =
  'CONCATENATE_OP'
  | 'EQUALS_OP'
  | 'NOT_EQUAL_OP'
  | 'GREATER_THAN_OP'
  | 'LESS_THAN_OP'
  | 'GREATER_THAN_OR_EQUAL_OP'
  | 'LESS_THAN_OR_EQUAL_OP'
  | 'PLUS_OP'
  | 'MINUS_OP'
  | 'TIMES_OP'
  | 'DIV_OP'
  | 'POWER_OP'

export interface FormulaBinaryOpAstNode extends FormulaAstNodeBase {
  type: FormulaBinaryOpType,
  left: FormulaAstNode,
  right: FormulaAstNode,
}

export type FormulaUnaryOpType =
  'MINUS_UNARY_OP'
  | 'PLUS_UNARY_OP'
  | 'PERCENT_OP'

export interface FormulaUnaryOpAstNode extends FormulaAstNodeBase {
  type: FormulaUnaryOpType,
  value: FormulaAstNode,
}

export interface FormulaFunctionCallAstNode extends FormulaAstNodeBase {
  type: 'FUNCTION_CALL',
  procedureName: string,
  args: FormulaAstNode[],
  internalWhitespace?: string,
  hyperlink?: string,
}

export interface FormulaNamedExpressionAstNode extends FormulaAstNodeBase {
  type: 'NAMED_EXPRESSION',
  expressionName: string,
}

export interface FormulaParenthesisAstNode extends FormulaAstNodeBase {
  type: 'PARENTHESES',
  expression: FormulaAstNode,
  internalWhitespace?: string,
}

export interface FormulaErrorAstNode extends FormulaAstNodeBase {
  type: 'ERROR',
  error: FormulaCellError,
}

export interface FormulaErrorWithRawInputAstNode extends FormulaAstNodeBase {
  type: 'ERROR_WITH_RAW_INPUT',
  rawInput: string,
  error: FormulaCellError,
}

export interface FormulaArrayAstNode extends FormulaAstNodeBase {
  type: 'ARRAY',
  args: FormulaAstNode[][],
  internalWhitespace?: string,
}

export interface FormulaCellError {
  type: string,
  message?: string,
}

export type FormulaCellReferenceType =
  'CELL_REFERENCE'
  | 'CELL_REFERENCE_ABSOLUTE'
  | 'CELL_REFERENCE_ABSOLUTE_COL'
  | 'CELL_REFERENCE_ABSOLUTE_ROW'

export interface FormulaCellReference {
  sheet: number,
  col: number,
  row: number,
  referenceType: FormulaCellReferenceType,
  sheetAbsolute: boolean,
  colAbsolute: boolean,
  rowAbsolute: boolean,
}

export interface FormulaColumnReference {
  sheet: number,
  col: number,
  sheetAbsolute: boolean,
  colAbsolute: boolean,
}

export interface FormulaRowReference {
  sheet: number,
  row: number,
  sheetAbsolute: boolean,
  rowAbsolute: boolean,
}

export type FormulaRangeSheetReferenceType =
  'RELATIVE'
  | 'START_ABSOLUTE'
  | 'BOTH_ABSOLUTE'

export type FormulaDependency =
  FormulaCellDependency
  | FormulaCellRangeDependency
  | FormulaColumnRangeDependency
  | FormulaRowRangeDependency
  | FormulaNamedExpressionDependency

export interface FormulaCellDependency {
  type: 'CELL',
  address: SimpleCellAddress,
}

export interface FormulaCellRangeDependency {
  type: 'CELL_RANGE',
  start: SimpleCellAddress,
  end: SimpleCellAddress,
}

export interface FormulaColumnRangeDependency {
  type: 'COLUMN_RANGE',
  sheet: number,
  startColumn: number,
  endColumn: number,
}

export interface FormulaRowRangeDependency {
  type: 'ROW_RANGE',
  sheet: number,
  startRow: number,
  endRow: number,
}

export interface FormulaNamedExpressionDependency {
  type: 'NAMED_EXPRESSION',
  name: string,
}

/**
 * @internal
 */
export const mapFormulaTokenizationResult = (result: ILexingResult): FormulaTokenizationResult => ({
  tokens: result.tokens.map(mapFormulaToken),
  errors: result.errors.map(mapFormulaLexingError),
})

/**
 * @internal
 */
export const mapFormulaParsingResult = (result: ParsingResult, formulaAddress: SimpleCellAddress): FormulaParsingResult => ({
  ast: result.errors.length === 0 ? mapFormulaAst(result.ast, formulaAddress) : undefined,
  errors: result.errors.map(mapFormulaParsingError),
  dependencies: result.dependencies.map(dependency => mapFormulaDependency(dependency, formulaAddress)),
  hasVolatileFunction: result.hasVolatileFunction,
  hasStructuralChangeFunction: result.hasStructuralChangeFunction,
})

const mapFormulaToken = (token: IToken): FormulaToken => ({
  type: token.tokenType.name,
  image: token.image,
  startOffset: token.startOffset,
  endOffset: token.endOffset ?? token.startOffset + token.image.length - 1,
  startLine: token.startLine,
  endLine: token.endLine,
  startColumn: token.startColumn,
  endColumn: token.endColumn,
})

const mapFormulaLexingError = (error: ILexingError): FormulaLexingError => ({
  message: error.message,
  offset: error.offset,
  line: error.line,
  column: error.column,
  length: error.length,
})

const mapFormulaParsingError = (error: ParsingError): FormulaParsingError => ({
  type: error.type,
  message: error.message,
})

const mapFormulaAst = (ast: Ast, formulaAddress: SimpleCellAddress): FormulaAstNode => {
  switch (ast.type) {
    case AstNodeType.EMPTY:
      return withLeadingWhitespace({type: 'EMPTY'}, ast.leadingWhitespace)
    case AstNodeType.NUMBER:
      return withLeadingWhitespace({type: 'NUMBER', value: ast.value}, ast.leadingWhitespace)
    case AstNodeType.STRING:
      return withLeadingWhitespace({type: 'STRING', value: ast.value}, ast.leadingWhitespace)
    case AstNodeType.CELL_REFERENCE:
      return withLeadingWhitespace({
        type: 'CELL_REFERENCE',
        reference: mapFormulaCellReference(ast.reference, formulaAddress),
      }, ast.leadingWhitespace)
    case AstNodeType.CELL_RANGE:
      return withLeadingWhitespace({
        type: 'CELL_RANGE',
        start: mapFormulaCellReference(ast.start, formulaAddress),
        end: mapFormulaCellReference(ast.end, formulaAddress),
        sheetReferenceType: mapFormulaRangeSheetReferenceType(ast.sheetReferenceType),
      }, ast.leadingWhitespace)
    case AstNodeType.COLUMN_RANGE:
      return withLeadingWhitespace({
        type: 'COLUMN_RANGE',
        start: mapFormulaColumnReference(ast.start, formulaAddress),
        end: mapFormulaColumnReference(ast.end, formulaAddress),
        sheetReferenceType: mapFormulaRangeSheetReferenceType(ast.sheetReferenceType),
      }, ast.leadingWhitespace)
    case AstNodeType.ROW_RANGE:
      return withLeadingWhitespace({
        type: 'ROW_RANGE',
        start: mapFormulaRowReference(ast.start, formulaAddress),
        end: mapFormulaRowReference(ast.end, formulaAddress),
        sheetReferenceType: mapFormulaRangeSheetReferenceType(ast.sheetReferenceType),
      }, ast.leadingWhitespace)
    case AstNodeType.CONCATENATE_OP:
    case AstNodeType.EQUALS_OP:
    case AstNodeType.NOT_EQUAL_OP:
    case AstNodeType.GREATER_THAN_OP:
    case AstNodeType.LESS_THAN_OP:
    case AstNodeType.GREATER_THAN_OR_EQUAL_OP:
    case AstNodeType.LESS_THAN_OR_EQUAL_OP:
    case AstNodeType.PLUS_OP:
    case AstNodeType.MINUS_OP:
    case AstNodeType.TIMES_OP:
    case AstNodeType.DIV_OP:
    case AstNodeType.POWER_OP:
      return withLeadingWhitespace({
        type: ast.type as FormulaBinaryOpType,
        left: mapFormulaAst(ast.left, formulaAddress),
        right: mapFormulaAst(ast.right, formulaAddress),
      }, ast.leadingWhitespace)
    case AstNodeType.MINUS_UNARY_OP:
    case AstNodeType.PLUS_UNARY_OP:
    case AstNodeType.PERCENT_OP:
      return withLeadingWhitespace({
        type: ast.type as FormulaUnaryOpType,
        value: mapFormulaAst(ast.value, formulaAddress),
      }, ast.leadingWhitespace)
    case AstNodeType.FUNCTION_CALL:
      return withInternalWhitespace(withLeadingWhitespace({
        type: 'FUNCTION_CALL',
        procedureName: ast.procedureName,
        args: ast.args.map(arg => mapFormulaAst(arg, formulaAddress)),
        hyperlink: ast.hyperlink,
      }, ast.leadingWhitespace), ast.internalWhitespace)
    case AstNodeType.NAMED_EXPRESSION:
      return withLeadingWhitespace({
        type: 'NAMED_EXPRESSION',
        expressionName: ast.expressionName,
      }, ast.leadingWhitespace)
    case AstNodeType.PARENTHESIS:
      return withInternalWhitespace(withLeadingWhitespace({
        type: 'PARENTHESES',
        expression: mapFormulaAst(ast.expression, formulaAddress),
      }, ast.leadingWhitespace), ast.internalWhitespace)
    case AstNodeType.ERROR:
      return withLeadingWhitespace({
        type: 'ERROR',
        error: mapFormulaCellError(ast.error),
      }, ast.leadingWhitespace)
    case AstNodeType.ERROR_WITH_RAW_INPUT:
      return withLeadingWhitespace({
        type: 'ERROR_WITH_RAW_INPUT',
        rawInput: ast.rawInput,
        error: mapFormulaCellError(ast.error),
      }, ast.leadingWhitespace)
    case AstNodeType.ARRAY:
      return withInternalWhitespace(withLeadingWhitespace({
        type: 'ARRAY',
        args: ast.args.map(row => row.map(arg => mapFormulaAst(arg, formulaAddress))),
      }, ast.leadingWhitespace), ast.internalWhitespace)
  }
}

const mapFormulaDependency = (dependency: RelativeDependency, formulaAddress: SimpleCellAddress): FormulaDependency => {
  if (dependency instanceof AddressDependency) {
    return {
      type: 'CELL',
      address: cloneSimpleCellAddress(dependency.dependency.toSimpleCellAddress(formulaAddress)),
    }
  } else if (dependency instanceof CellRangeDependency) {
    return {
      type: 'CELL_RANGE',
      start: cloneSimpleCellAddress(dependency.start.toSimpleCellAddress(formulaAddress)),
      end: cloneSimpleCellAddress(dependency.end.toSimpleCellAddress(formulaAddress)),
    }
  } else if (dependency instanceof ColumnRangeDependency) {
    const start = dependency.start.toSimpleColumnAddress(formulaAddress)
    const end = dependency.end.toSimpleColumnAddress(formulaAddress)
    return {
      type: 'COLUMN_RANGE',
      sheet: start.sheet,
      startColumn: start.col,
      endColumn: end.col,
    }
  } else if (dependency instanceof RowRangeDependency) {
    const start = dependency.start.toSimpleRowAddress(formulaAddress)
    const end = dependency.end.toSimpleRowAddress(formulaAddress)
    return {
      type: 'ROW_RANGE',
      sheet: start.sheet,
      startRow: start.row,
      endRow: end.row,
    }
  } else if (dependency instanceof NamedExpressionDependency) {
    return {
      type: 'NAMED_EXPRESSION',
      name: dependency.name,
    }
  } else {
    throw new Error('Unknown formula dependency type')
  }
}

const mapFormulaCellReference = (address: CellAddress, formulaAddress: SimpleCellAddress): FormulaCellReference => {
  const simpleAddress = address.toSimpleCellAddress(formulaAddress)
  const colAbsolute = address.isColumnAbsolute()
  const rowAbsolute = address.isRowAbsolute()

  return {
    sheet: simpleAddress.sheet,
    col: simpleAddress.col,
    row: simpleAddress.row,
    referenceType: mapFormulaCellReferenceType(colAbsolute, rowAbsolute),
    sheetAbsolute: address.sheet !== undefined,
    colAbsolute,
    rowAbsolute,
  }
}

const mapFormulaColumnReference = (address: ColumnAddress, formulaAddress: SimpleCellAddress): FormulaColumnReference => {
  const simpleAddress = address.toSimpleColumnAddress(formulaAddress)

  return {
    sheet: simpleAddress.sheet,
    col: simpleAddress.col,
    sheetAbsolute: address.sheet !== undefined,
    colAbsolute: address.isColumnAbsolute(),
  }
}

const mapFormulaRowReference = (address: RowAddress, formulaAddress: SimpleCellAddress): FormulaRowReference => {
  const simpleAddress = address.toSimpleRowAddress(formulaAddress)

  return {
    sheet: simpleAddress.sheet,
    row: simpleAddress.row,
    sheetAbsolute: address.sheet !== undefined,
    rowAbsolute: address.isRowAbsolute(),
  }
}

const mapFormulaCellReferenceType = (colAbsolute: boolean, rowAbsolute: boolean): FormulaCellReferenceType => {
  if (colAbsolute && rowAbsolute) {
    return 'CELL_REFERENCE_ABSOLUTE'
  } else if (colAbsolute) {
    return 'CELL_REFERENCE_ABSOLUTE_COL'
  } else if (rowAbsolute) {
    return 'CELL_REFERENCE_ABSOLUTE_ROW'
  } else {
    return 'CELL_REFERENCE'
  }
}

const mapFormulaRangeSheetReferenceType = (sheetReferenceType: RangeSheetReferenceType): FormulaRangeSheetReferenceType => {
  switch (sheetReferenceType) {
    case RangeSheetReferenceType.RELATIVE:
      return 'RELATIVE'
    case RangeSheetReferenceType.START_ABSOLUTE:
      return 'START_ABSOLUTE'
    case RangeSheetReferenceType.BOTH_ABSOLUTE:
      return 'BOTH_ABSOLUTE'
    default:
      throw new Error('Unknown formula range sheet reference type')
  }
}

const mapFormulaCellError = (error: {type: string, message?: string}): FormulaCellError => ({
  type: error.type,
  message: error.message,
})

const withLeadingWhitespace = <T extends FormulaAstNodeBase>(node: T, leadingWhitespace?: string): T => {
  if (leadingWhitespace !== undefined) {
    node.leadingWhitespace = leadingWhitespace
  }

  return node
}

const withInternalWhitespace = <T extends FormulaAstNodeBase & {internalWhitespace?: string}>(node: T, internalWhitespace?: string): T => {
  if (internalWhitespace !== undefined) {
    node.internalWhitespace = internalWhitespace
  }

  return node
}

const cloneSimpleCellAddress = (address: SimpleCellAddress): SimpleCellAddress => ({
  sheet: address.sheet,
  col: address.col,
  row: address.row,
})
