/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { items } from '#common/linq';
import { KnownMediaType } from '#common/media-types';
import { Expression, ExpressionOrLiteral } from '#csharp/code-dom/expression';
import { Interface } from '#csharp/code-dom/interface';
import { Namespace } from '#csharp/code-dom/namespace';
import { OneOrMoreStatements } from '#csharp/code-dom/statements/statement';
import { Variable } from '#csharp/code-dom/variable';
import { ClientRuntime } from '#csharp/lowlevel-generator/clientruntime';
import { Schema } from '#csharp/lowlevel-generator/code-model';
import { ModelClass } from '#csharp/lowlevel-generator/model/model-class';
import { EnhancedTypeDeclaration } from '#csharp/schema/extended-type-declaration';
import { State } from '../generator';
import { ModelInterfaceProperty } from './interface-property';

export class ModelInterface extends Interface implements EnhancedTypeDeclaration {
  get schema(): Schema {
    return this.implementation.schema;
  }
  deserializeFromContainerMember(mediaType: KnownMediaType, container: ExpressionOrLiteral, serializedName: string, defaultValue: Expression): Expression {
    return this.implementation.deserializeFromContainerMember(mediaType, container, serializedName, defaultValue);
  }
  deserializeFromNode(mediaType: KnownMediaType, node: ExpressionOrLiteral, defaultValue: Expression): Expression {
    return this.implementation.deserializeFromNode(mediaType, node, defaultValue);
  }
  /** emits an expression to deserialize content from a string */
  deserializeFromString(mediaType: KnownMediaType, content: ExpressionOrLiteral, defaultValue: Expression): Expression | undefined {
    return this.implementation.deserializeFromString(mediaType, content, defaultValue);
  }

  /** emits an expression to deserialize content from a content/response */
  deserializeFromResponse(mediaType: KnownMediaType, content: ExpressionOrLiteral, defaultValue: Expression): Expression | undefined {
    return this.implementation.deserializeFromResponse(mediaType, content, defaultValue);
  }


  /** emits an expression serialize this to a HttpContent */
  serializeToContent(mediaType: KnownMediaType, value: ExpressionOrLiteral): Expression {
    return this.implementation.serializeToContent(mediaType, value);
  }

  serializeToNode(mediaType: KnownMediaType, value: ExpressionOrLiteral, serializedName: string): Expression {
    return this.implementation.serializeToNode(mediaType, value, serializedName);
  }
  serializeToContainerMember(mediaType: KnownMediaType, value: ExpressionOrLiteral, container: Variable, serializedName: string): OneOrMoreStatements {
    return this.implementation.serializeToContainerMember(mediaType, value, container, serializedName);
  }

  get isXmlAttribute(): boolean {
    return this.implementation.isXmlAttribute;
  }

  get isRequired(): boolean {
    return this.implementation.isRequired;
  }

  public validatePresence(eventListener: Variable, property: Variable): OneOrMoreStatements {
    return this.implementation.validatePresence(eventListener, property);
  }
  public validateValue(eventListener: Variable, property: Variable): OneOrMoreStatements {
    return this.implementation.validateValue(eventListener, property);
  }


  get hasHeaderProperties(): boolean {
    return this.implementation.hasHeaderProperties;
  }
  constructor(parent: Namespace, schema: Schema, public implementation: ModelClass, public state: State, objectInitializer?: Partial<ModelInterface>) {
    super(parent, `I${schema.details.csharp.name}`);
    this.partial = true;
    this.apply(objectInitializer);
    const implData = (schema.details.csharp = schema.details.csharp || {});
    implData.interfaceImplementation = this;
    this.description = `${schema.details.default.description}`;

    for (const { key: propertyName, value: property } of items(schema.properties)) {
      this.add(new ModelInterfaceProperty(this, property, state.path('properties', propertyName)));
    }

    // mark it as json serializable
    if (!schema.details.default.isHeaderModel) {
      if (this.state.project.jsonSerialization) {
        this.interfaces.push(ClientRuntime.IJsonSerializable);
      }
      if (this.state.project.xmlSerialization) {
        this.interfaces.push(ClientRuntime.IXmlSerializable);
      }
    }

  }
}
