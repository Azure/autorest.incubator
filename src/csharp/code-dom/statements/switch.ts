/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Initializer } from '#common/initializer';
import { EOL, indent } from '#common/text-manipulation';
import { Expression, ExpressionOrLiteral, toExpression } from '#csharp/code-dom/expression';
import { CaseStatement } from '#csharp/code-dom/statements/case';
import { Statement } from '#csharp/code-dom/statements/statement';

export type OneOrMoreCases = (() => Iterable<CaseStatement>) | Iterable<CaseStatement> | CaseStatement;

export function Switch(expression: ExpressionOrLiteral, cases: OneOrMoreCases, objectInitializer?: Partial<SwitchStatement>): SwitchStatement {
  return new SwitchStatement(toExpression(expression), cases, objectInitializer);
}

export class SwitchStatement extends Initializer implements Statement {
  protected caseStatements = new Array<CaseStatement>();

  public get implementation(): string {
    return `
switch ( ${this.expression.value} )
{
${indent(this.caseStatements.map(each => each.implementation).join(EOL))}
}`;
  }
  constructor(protected expression: Expression, cases: OneOrMoreCases, objectInitializer?: Partial<SwitchStatement>) {
    super();
    this.apply(objectInitializer);
    this.add(cases);
  }
  public add(cases: OneOrMoreCases) {
    if (typeof cases === 'function') {
      cases = cases();
    }
    if (cases instanceof CaseStatement) {
      this.caseStatements.push(cases);
    }
    else {
      for (const caseStatement of cases) {
        this.caseStatements.push(caseStatement);
      }
    }
  }
}