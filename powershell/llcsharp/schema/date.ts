/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { StringExpression } from '@azure-tools/codegen-csharp';
import { Schema } from '../code-model';
import { DateSchema } from '@azure-tools/codemodel';
import { NewDateTime } from './date-time';


export class NewDate extends NewDateTime {
  public DateTimeFormat = new StringExpression('yyyy-MM-dd');
  constructor(schema: DateSchema, isRequired: boolean) {
    super(schema, isRequired);
  }
}