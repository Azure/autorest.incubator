/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { values } from '#common/linq';
import * as aio from '@microsoft.azure/async-io';
import { join } from 'path';

function getAllPropertyNames(obj: any) {
  const props = new Array<string>();

  do {
    Object.getOwnPropertyNames(obj).forEach(function (prop) {
      if (props.indexOf(prop) === -1) {
        props.push(prop);
      }
    });
  } while (obj = Object.getPrototypeOf(obj));

  return props;
}

export function fail(text: string): never {
  throw new Error(text);
}

export async function copyResources(sourceFolder: string, fileWriter: (filename: string, content: string) => Promise<void>) {
  const done = new Array<Promise<void>>();
  try {
    const files = await aio.readdir(sourceFolder);

    for (const file of values(files)) {
      const fullPath = join(sourceFolder, file);
      if (await aio.isDirectory(fullPath)) {
        done.push(copyResources(fullPath, async (f, c) => fileWriter(`${file}/${f}`, c)));
        continue;
      }
      if (await aio.isFile(fullPath)) {
        done.push(aio.readFile(fullPath).then(async (content) => fileWriter(file, content)));
      }
    }
  } catch {

  }
  await Promise.all(done);
}
