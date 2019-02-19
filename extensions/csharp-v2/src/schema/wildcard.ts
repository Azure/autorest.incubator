/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { KnownMediaType } from '@microsoft.azure/autorest.codemodel-v3';
import { nameof, camelCase, deconstruct } from '@microsoft.azure/codegen';
import { Expression, ExpressionOrLiteral, toExpression, valueOf, IsNotNull, And } from '@microsoft.azure/codegen-csharp';
import { ForEach } from '@microsoft.azure/codegen-csharp';
import { If } from '@microsoft.azure/codegen-csharp';
import { OneOrMoreStatements } from '@microsoft.azure/codegen-csharp';
import { Variable } from '@microsoft.azure/codegen-csharp';
import { Schema } from '../code-model';
import { popTempVar, pushTempVar } from './primitive';
import { EnhancedTypeDeclaration } from './extended-type-declaration';
import { ClientRuntime } from '../clientruntime';
import { System } from '@microsoft.azure/codegen-csharp';
import { TypeDeclaration } from '@microsoft.azure/codegen-csharp';

export class Wildcard implements EnhancedTypeDeclaration {
  public isXmlAttribute: boolean = false;

  constructor(public schema: Schema, protected leafType: EnhancedTypeDeclaration) {
  }

  get declaration(): string {
    return `System.Collections.Hashtable`;
  }

  /** emits an expression to deserialize a property from a member inside a container */
  deserializeFromContainerMember(mediaType: KnownMediaType, container: ExpressionOrLiteral, serializedName: string, defaultValue: Expression): Expression {
    switch (mediaType) {
      case KnownMediaType.Json:
        return toExpression(`/* 1 */ new System.Collections.Hashtable( System.Linq.Enumerable.ToDictionary<string,string, ${this.leafType.declaration}>( ${valueOf(container)}.Property("${serializedName}")?.Keys ?? System.Linq.Enumerable.Empty<string>(), each => each, each => ${this.leafType.deserializeFromNode(mediaType, `${valueOf(container)}.Property("${serializedName}").PropertyT<${ClientRuntime.JsonNode}>(each)`, toExpression('null'))} ))`);

      case KnownMediaType.Xml:
        return toExpression(`new System.Collections.Hashtable(System.Linq.Enumerable.ToDictionary<string,string, ${this.leafType.declaration}>( ${valueOf(container)}?.Elements() ?? System.Linq.Enumerable.Empty<System.Xml.Linq.XElement>(), element => element.Name.ToString(), element => ${this.leafType.deserializeFromNode(mediaType, 'element', toExpression('null'))} ))`);

      case KnownMediaType.Header: {
        const prefix = this.schema.extensions['x-ms-header-collection-prefix'];
        if (prefix) {
          // this is a catch for a specific set of headers
          return toExpression(`new System.Collections.Hashtable(System.Linq.Enumerable.ToDictionary<string,string, ${this.leafType.declaration}>(System.Linq.Enumerable.Where(${valueOf(container)}, header => header.Key.StartsWith("${serializedName}")), header => header.Key.Substring(${serializedName.length}), header => System.Linq.Enumerable.FirstOrDefault(header.Value)))`);
        }
        return toExpression(`new System.Collections.Hashtable(System.Linq.Enumerable.ToDictionary<string,string, ${this.leafType.declaration}>( ${valueOf(container)}?.Elements() ?? System.Linq.Enumerable.Empty<System.Xml.Linq.XElement>(), element => element.Name.ToString(), element => ${this.leafType.deserializeFromNode(mediaType, 'element', toExpression('null'))} ))`);
      }
    }
    return toExpression(`null /* deserializeFromContainerMember (wildcard) doesn't support '${mediaType}' ${__filename}*/`);
  }

  /** emits an expression to deserialze a container as the value itself. */
  deserializeFromNode(mediaType: KnownMediaType, node: ExpressionOrLiteral, defaultValue: Expression): Expression {
    switch (mediaType) {
      case KnownMediaType.Json:
        const nodeAsObject = `(${(valueOf(node))} as ${ClientRuntime.JsonObject})`;
        return toExpression(`/* 2 */ new System.Collections.Hashtable(System.Linq.Enumerable.ToDictionary<string,string, ${this.leafType.declaration}>( ${nodeAsObject}?.Keys ?? System.Linq.Enumerable.Empty<string>(), each => each, each => ${this.leafType.deserializeFromNode(mediaType, `${nodeAsObject}.PropertyT<${ClientRuntime.JsonNode}>(each)`, toExpression('null'))} ))`);
    }
    return toExpression(`null /* deserializeFromNode (wildcard) doesn't support '${mediaType}' ${__filename}*/`);
  }

  /** emits an expression serialize this to the value required by the container */
  serializeToNode(mediaType: KnownMediaType, value: ExpressionOrLiteral, serializedName: string): Expression {
    try {
      const each = pushTempVar();

      switch (mediaType) {
        case KnownMediaType.Json: {
          const serDict = ` System.Linq.Enumerable.Select( (System.Collections.Generic.IEnumerable<object>)${value}.Keys, (${each}) => new System.Collections.Generic.KeyValuePair<${System.String}, ${ClientRuntime.JsonNode}>( ${each}.ToString(), ${this.leafType.serializeToNode(mediaType, `${value}[${each}]`, serializedName)}))`;
          return toExpression(`null != ${value} ? new ${ClientRuntime.JsonObject}(${serDict}) : null`);
        }

      }
      return toExpression(`null /* serializeToNode (wildcard) doesn't support '${mediaType}' ${__filename}*/`);
    } finally {
      popTempVar();
    }
  }
  /** emits an expression serialize this to the value required by the container */
  serializeToContent(mediaType: KnownMediaType, value: ExpressionOrLiteral): Expression {
    return toExpression(`null /* serializeToContent (wildcard) doesn't support '${mediaType}' ${__filename}*/`);
  }
  /** emits an expression to deserialize content from a string */
  deserializeFromString(mediaType: KnownMediaType, content: ExpressionOrLiteral, defaultValue: Expression): Expression | undefined {
    return toExpression(`null /* deserializeFromString (wildcard)doesn't support '${mediaType}' ${__filename}*/`);
  }
  /** emits an expression to deserialize content from a content/response */
  deserializeFromResponse(mediaType: KnownMediaType, content: ExpressionOrLiteral, defaultValue: Expression): Expression | undefined {
    return toExpression(`null /* deserializeFromResponse doesn't support '${mediaType}' ${__filename}*/`);
  }

  /** emits the code required to serialize this into a container */
  serializeToContainerMember(mediaType: KnownMediaType, value: ExpressionOrLiteral, container: Variable, serializedName: string): OneOrMoreStatements {
    const $this = this;
    try {
      const each = pushTempVar();
      const item = pushTempVar();
      const eachvalue = pushTempVar();
      const innerContainer = pushTempVar();

      switch (mediaType) {
        case KnownMediaType.Json:
          return If(`null != ${value}`, function* () {
            yield `var ${innerContainer} = ${ClientRuntime.JsonObject.new()};`;
            yield `${container}.Add("${serializedName}", ${innerContainer});`;
            yield ForEach(each, `${toExpression(value)}.Keys`, function* () {
              const eachv = `${value}[${each}]`;

              yield If(And(IsNotNull(eachv), `${eachv} is ${$this.leafType.declaration} ${eachvalue}`),
                `AddIf( ${$this.leafType.serializeToNode(mediaType, `${eachvalue}`, `$$$`)},(${item}) => ${innerContainer}.Add(${each} as string,${item} ) );`);

              //` ${eachvalue} = ${value}[${each}];`,  );
              //yield `AddIf( ${$this.leafType.serializeToNode(mediaType, `(${eachvalue} as ${$this.leafType.declaration})`, `$$$`)},(${item}) => ${innerContainer}.Add(${each} as string,${item} ) );`;
            });
          });

        case KnownMediaType.Xml:
          return If(`null != ${value}`, ForEach(each, toExpression(value), `AddIf( ${this.leafType.serializeToNode(mediaType, `${each}.Value`, `$$$`)},${container}.Add );`.replace('"$$$"', `${each}.Key`)));

        case KnownMediaType.Header:
          const prefix = this.schema.extensions['x-ms-header-collection-prefix'];
          if (prefix) {
            return If(`null != ${value}`, ForEach(each, toExpression(value), `${valueOf(container)}.Add("${prefix}"+${each}.Key,${each}.Value);`));
          }
      }
    } finally {
      popTempVar();
      popTempVar();
      popTempVar();
      popTempVar();
    }
    return `/* serializeToContainerMember (wildcard) doesn't support '${mediaType}' ${__filename}*/`;
  }
  /** is the value of this type permitted to be NULL */
  isRequired = false;

  public validatePresence(eventListener: Variable, property: Variable): string {
    return ``;
  }
  validateValue(eventListener: Variable, property: Variable): string {
    return ``;
  }
}

export class UntypedWildcard implements EnhancedTypeDeclaration {
  public isXmlAttribute: boolean = false;

  constructor(public schema: Schema) {
  }

  static deserializeFromContainerMemberToType(mediaType: KnownMediaType, container: ExpressionOrLiteral, serializedName: string, defaultValue: Expression, typeDeclaration: TypeDeclaration): Expression {
    try {
      switch (mediaType) {
        case KnownMediaType.Json:
          const tmp = `__${camelCase(['json', ...deconstruct(serializedName)])}`;
          return toExpression(`If( ${valueOf(container)}?.PropertyT<${ClientRuntime.JsonObject}>("${serializedName}"), out var ${tmp}) ? ${ClientRuntime.JsonSerializable}.FromJson(${tmp}, ${defaultValue} ?? new System.Collections.Hashtable( new ${typeDeclaration.declaration}())) : ${defaultValue}`);
      }
    } finally {
    }

    return toExpression(`null /* deserializeFromContainerMemberToType doesn't support '${mediaType}' ${__filename}*/`);
  }

  /** emits an expression to deserialize a property from a member inside a container */
  deserializeFromContainerMember(mediaType: KnownMediaType, container: ExpressionOrLiteral, serializedName: string, defaultValue: Expression): Expression {
    return UntypedWildcard.deserializeFromContainerMemberToType(mediaType, container, serializedName, defaultValue, System.Collections.Generic.Dictionary(System.String, System.Object));
  }

  /** emits an expression to deserialze a container as the value itself. */
  deserializeFromNode(mediaType: KnownMediaType, node: ExpressionOrLiteral, defaultValue: Expression): Expression {
    return toExpression(`null /* deserializeFromNode doesn't support '${mediaType}' ${__filename}*/`);
  }

  /** emits an expression serialize this to the value required by the container */
  serializeToNode(mediaType: KnownMediaType, value: ExpressionOrLiteral, serializedName: string): Expression {
    return toExpression(`null /* serializeToNode doesn't support '${mediaType}' ${__filename}*/`);
  }
  /** emits an expression serialize this to the value required by the container */
  serializeToContent(mediaType: KnownMediaType, value: ExpressionOrLiteral): Expression {
    return toExpression(`null /* serializeToContent doesn't support '${mediaType}' ${__filename}*/`);
  }
  /** emits an expression to deserialize content from a string */
  deserializeFromString(mediaType: KnownMediaType, content: ExpressionOrLiteral, defaultValue: Expression): Expression | undefined {
    return toExpression(`null /* deserializeFromString doesn't support '${mediaType}' ${__filename}*/`);
  }

  /** emits an expression to deserialize content from a content/response */
  deserializeFromResponse(mediaType: KnownMediaType, content: ExpressionOrLiteral, defaultValue: Expression): Expression | undefined {
    return toExpression(`null /* deserializeFromResponse doesn't support '${mediaType}' ${__filename}*/`);
  }

  /** emits the code required to serialize this into a container */
  serializeToContainerMember(mediaType: KnownMediaType, value: ExpressionOrLiteral, container: Variable, serializedName: string): OneOrMoreStatements {
    switch (mediaType) {
      case KnownMediaType.Json:
        return If(`${value}.Count > 0`, function* () {
          yield `${container}.Add("${serializedName}", ${ClientRuntime.JsonSerializable}.ToJsonValue(${value}));`;
        });

    }
    return `/* serializeToContainerMember doesn't support '${mediaType}' ${__filename}*/`;
  }

  /** is the value of this type permitted to be NULL */
  isRequired = false;

  get declaration(): string {
    return `System.Collections.Hashtable`;
  }
  public validatePresence(eventListener: Variable, property: Variable): string {
    return `await ${eventListener}.AssertNotNull(${nameof(property.value)},${property});`.trim();
  }
  validateValue(eventListener: Variable, property: Variable): string {
    return `/* untyped wildcard validate value for ${property} */`;
  }
}
