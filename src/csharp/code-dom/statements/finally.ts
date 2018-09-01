/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { indent } from '#common/text-manipulation';
import { StatementPossibilities, Statements } from '#csharp/code-dom/statements/statement';

export function Finally(body: StatementPossibilities, objectInitializer?: Partial<FinallyStatement>): FinallyStatement {
  return new FinallyStatement(body, objectInitializer);
}

export class FinallyStatement extends Statements {
  constructor(body: StatementPossibilities, objectInitializer?: Partial<FinallyStatement>) {
    super(body);
    this.apply(objectInitializer);
  }

  public get implementation(): string {
    return `
finally
{
${indent(super.implementation)}
} `.trim();
  }
}