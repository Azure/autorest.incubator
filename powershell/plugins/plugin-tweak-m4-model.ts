/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
import { ArraySchema, CodeModel, DictionarySchema, getAllProperties, HttpHeader, ObjectSchema, Property, Schema, SchemaType } from '@azure-tools/codemodel';
import { serialize } from '@azure-tools/codegen';
import { PwshModel } from '../utils/PwshModel';
import { ModelState } from '../utils/model-state';
import { StatusCodes } from '../utils/http-definitions';
import { items, values, keys, Dictionary, length } from '@azure-tools/linq';

import { Host } from '@azure-tools/autorest-extension-base';

type State = ModelState<PwshModel>;

let directives: Array<any> = [];

async function tweakModel(state: State): Promise<PwshModel> {
  const model = state.model;

  addResponseHeaderSchema(model);

  addDictionaryApiVersion(model);

  removeM4DefaultDescription(model);

  handleNoinlineDirective(state);

  handleDiscriminator(state);

  return model;
}

function handleNoinlineDirective(state: State) {
  var inlineModels: string[] = new Array;
  for (const directive of directives.filter(each => each['no-inline'])) {
    inlineModels = inlineModels.concat(<ConcatArray<string>>values(directive['no-inline']).toArray());
  }
  for (const model of state.model.schemas.objects || []) {
    if (inlineModels.includes(model.language.default.name)) {
      model.language.default['skip-inline'] = true;
    }
  }
}

function handleDiscriminator(state: State) {
  for (const model of state.model.schemas.objects || []) {
    if (model.discriminator) {
      model.language.default['skip-inline'] = true;
    }
  }
}

function addResponseHeaderSchema(model: CodeModel) {
  // In remodeler, each operations response headers will has its own scheam. Each header will be schema's property. 
  // But in m4, if 'schema' is not explicitly defined, even 'headers' is specified, there won't be a schema for headers.
  // To keep backward compatiable, we will create headers schema here

  model.operationGroups.forEach((group) => {
    group.operations?.forEach((op) => {
      if (!op.responses) {
        return;
      }
      op.responses.forEach((resp) => {
        if ((<any>resp).schema) {
          return;
        }

        const headers = resp.protocol.http?.headers as Array<HttpHeader>;
        if (headers === undefined) {
          return;
        }

        const responseCode = resp.protocol.http?.statusCodes?.[0];
        if (responseCode === undefined) {
          return;
        }

        // Follow naming pattern in m3
        const code = ((<any>StatusCodes)[responseCode] || '') || responseCode;
        const schemaName = `${group.language.default.name}_${op.language.default.name} ${code} ResponseHeaders`;

        const newSchema = model.schemas.objects?.find((schema) => schema.language.default.name === schemaName) ||
          new ObjectSchema(schemaName, '');
        newSchema.language.default.isHeaderModel = true;

        if (!model.schemas.objects) {
          model.schemas.objects = [];
        }
        model.schemas.objects.push(newSchema);

        headers.forEach((head) => {
          // We lost description and x-ms-client-name in m4. So newProp's description is empty and use header as serializedName 
          const newProp = new Property(head.header, '', head.schema, {
            readOnly: false,
            required: false,
            serializedName: head.header
          });
          newProp.language.default.HeaderProperty = 'Header';

          if (!newSchema.properties) {
            newSchema.properties = [];
          }
          newSchema.properties.push(newProp);
        });

        // Set response header use new schema
        resp.language.default.headerSchema = newSchema;
      });
    });
  });
}

function addDictionaryApiVersion(model: CodeModel) {

  model.schemas.dictionaries?.forEach((schema) => {
    if (schema.apiVersions) {
      return;
    }
    if (schema.elementType && schema.elementType.apiVersions) {
      schema.apiVersions = JSON.parse(JSON.stringify(schema.elementType.apiVersions));
    }
  })

  // If we cannot find api version from element type, try to get it from object schema who refers the dict or any.

  model.schemas.objects?.forEach((schema) => {
    if (!schema.apiVersions) {
      return;
    }

    for (const prop of getAllProperties(schema)) {
      if (prop.schema.type !== SchemaType.Dictionary || prop.schema.apiVersions) {
        continue;
      }
      prop.schema.apiVersions = JSON.parse(JSON.stringify(schema.apiVersions));
    }
  })
}

function removeM4DefaultDescription(model: CodeModel) {
  // For dictionary and arrya schema and property, if there is no description assigned, m4 will set a default description like: Dictionary of <type> or Array of <type>
  // To keep same action as m3, we will set it to empty string

  const visited = new Set<Schema>();
  [...model.schemas.objects ?? [], ...model.schemas.dictionaries ?? [], ...model.schemas.arrays ?? []].forEach((schema) => {
    recursiveRemoveM4DefaultDescription(schema, visited);
  })
}

function recursiveRemoveM4DefaultDescription(schema: Schema, visited: Set<Schema>) {
  if (visited.has(schema) || (schema.type !== SchemaType.Object && schema.type !== SchemaType.Dictionary && schema.type !== SchemaType.Array)) {
    return;
  }
  // Default description pattern in m4
  const defaultDictDescPattern = /Dictionary of <.?>$/;
  const defaultArrayDescPattern = /Array of .?$/;
  visited.add(schema);
  if (schema.type === SchemaType.Dictionary) {
    const dictSchema = schema as DictionarySchema;
    recursiveRemoveM4DefaultDescription(dictSchema.elementType, visited);
    if (defaultDictDescPattern.test(dictSchema.language.default.description)) {
      dictSchema.language.default.description = '';
    }
  } else if (schema.type === SchemaType.Array) {
    const arrSchema = schema as ArraySchema;
    recursiveRemoveM4DefaultDescription(arrSchema.elementType, visited);
    if (defaultArrayDescPattern.test(schema.language.default.description)) {
      schema.language.default.description = '';
    }
  } else if (schema.type === SchemaType.Object) {
    const objSchema = schema as ObjectSchema;
    for (const prop of getAllProperties(objSchema)) {
      recursiveRemoveM4DefaultDescription(prop.schema, visited);
      if (prop.schema.type === SchemaType.Dictionary && (defaultDictDescPattern.test(prop.language.default.description) || defaultArrayDescPattern.test(prop.language.default.description))) {
        prop.language.default.description = '';
      }
    }
  }
}

export async function tweakM4ModelPlugin(service: Host) {
  const allDirectives = await service.GetValue('directive');
  directives = values(allDirectives).toArray();
  const state = await new ModelState<PwshModel>(service).init();
  service.WriteFile('code-model-v4-tweakm4codemodel.yaml', serialize(await tweakModel(state)), undefined, 'code-model-v4');
}
