/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { indent } from '#common/text-manipulation';
import { Expression, valueOf } from '#csharp/code-dom/expression';
import { StatementPossibilities, Statements } from '#csharp/code-dom/statements/statement';

export function For(initialization: Expression, condition: Expression, loop: Expression, statements: StatementPossibilities, objectInitializer?: Partial<ForStatement>) {
  return new ForStatement(initialization, condition, loop, statements, objectInitializer);
}

export class ForStatement extends Statements {
  constructor(public initialization: Expression, public condition: Expression, public loop: Expression, statements: StatementPossibilities, objectInitializer?: Partial<ForStatement>) {
    super(statements);
    this.apply(objectInitializer);
  }
  public get implementation(): string {
    return `
for( ${valueOf(this.initialization)} ; ${valueOf(this.condition)} ; ${valueOf(this.loop)})
{
${indent(super.implementation)}
}`.trim();
  }
}

export function ForEach(variable: string, enumerable: Expression, statements: StatementPossibilities, objectInitializer?: Partial<ForStatement>) {
  return new ForEachStatement(variable, enumerable, statements, objectInitializer);
}

export class ForEachStatement extends Statements {
  constructor(public variable: string, public enumerable: Expression, statements: StatementPossibilities, objectInitializer?: Partial<ForStatement>) {
    super(statements);
    this.apply(objectInitializer);
  }
  public get implementation(): string {
    return `
foreach( var ${this.variable} in ${valueOf(this.enumerable)} )
{
${indent(super.implementation)}
}`.trim();
  }
}
