/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Model } from '#common/code-model/code-model';
import { Text } from '#common/file-generator';
import { resources } from '#common/locations';
import { copyResources } from '#common/utility';
import { deserialize, serialize } from '#common/yaml';
import { Host } from '@microsoft.azure/autorest-extension-base';
import { join } from 'path';
import { Project } from './project';
import { State } from './state';

export async function processRequest(service: Host) {
  try {
    // Get the list of files
    const files = await service.ListInputs('code-model-v2');
    if (files.length === 0) {
      throw new Error('Inputs missing.');
    }

    const codemodel = files[0];

    // get the openapi document
    const codeModelText = await service.ReadFile(codemodel);
    const model = await deserialize<Model>(codeModelText, codemodel);

    // generate some files
    const modelState = new State(service, model, codemodel);
    const project = await new Project(modelState).init();
    await project.writeFiles(async (filename, content) => service.WriteFile(filename, content, undefined, 'source-file-csharp'));

    await service.ProtectFiles(project.csproj);
    await service.ProtectFiles(project.customFolder);
    await service.ProtectFiles(project.testFolder);

    // wait for all the generation to be done
    await generateCsproj(service, project);
    await copyRuntime(service, project);
    await generateCsproj(service, project);
    await generateModule(service, project);
    // await generateProxies(service);

    // debug data
    service.WriteFile('code-model-v2.powershell.yaml', serialize(model), undefined, 'source-file-other');

  } catch (E) {
    console.error(E);
    throw E;
  }
}

/*
async function generateProxies(service: Host, project: Project) {
  const of = await service.GetValue('output-folder');
  // find the pwsh executable.
  const pwsh = await resolveFullPath('pwsh', process.platform === 'win32' ? ['c:/Program Files/PowerShell', 'c:/Program Files (x86)/PowerShell'] : []);
  if (!pwsh) {
    // no powershell core found.
    throw new Error('PowerShell Core (pwsh) not found in path. Please ensure that pwsh is available.');
  }
  console.error(`${pwsh} -command "${of}/generate_proxies.ps1"`);
}
*/

async function copyRuntime(service: Host, project: Project) {
  // PowerShell Scripts
  await copyResources(join(resources, 'scripts', 'powershell'), async (fname, content) => service.WriteFile(fname, content, undefined, 'source-file-csharp'));

  // c# files
  await copyResources(join(resources, 'runtime', 'powershell'), async (fname, content) => service.WriteFile(join(project.runtimefolder, fname), content, undefined, 'source-file-csharp'));
  if (project.azure) {
    await copyResources(join(resources, 'runtime', 'powershell.azure'), async (fname, content) => service.WriteFile(join(project.runtimefolder, fname), content, undefined, 'source-file-csharp'));
  }
}

async function generateCsproj(service: Host, project: Project) {
  // write out the csproj file if it's not there.
  if (!await service.ReadFile(project.csproj)) {
    service.WriteFile(project.csproj, `<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <LangVersion>7.1</LangVersion>
    <OutputType>Library</OutputType>
     <TargetFramework>netstandard2.0</TargetFramework>
    <nowarn>1998</nowarn> <!-- some methods are marked async that don't have an await in them.-->
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="System.Management.Automation.dll" Version="10.0.10586" />
    <PackageReference Include="Microsoft.CSharp" Version="4.4.1" />
    <PackageReference Include="System.Text.Encodings.Web" Version="4.3.0" />
  </ItemGroup>
</Project>
`, undefined, 'source-file-csharp');
  }
}

async function generateModule(service: Host, project: Project) {
  // write out the psd1 file if it's not there.

  // if (!await service.ReadFile(cfg.psd1)) {

  // todo: change this to *update* the psd1?

  service.WriteFile(project.psd1, new Text(function* () {
    yield `@{`;
    yield `ModuleVersion="1.0"`;
    yield `NestedModules = @(`;
    yield `  "./bin/${project.moduleName}.private.dll"`;
    yield `  "${project.psm1}"`;
    yield `)`;
    yield `# don't export any actual cmdlets by default`;
    yield `CmdletsToExport = ''`;

    yield `# export the functions that we loaded(these are the proxy cmdlets)`;
    yield `FunctionsToExport = '*-*'`;
    yield `}`;
  }).text, undefined, 'source-file-powershell');

  // write out the psm1 file if it's not there.

  const psm1 = new Text(await service.ReadFile(project.psm1) || '');

  // clear regions first
  psm1.removeRegion('Initialization');
  psm1.removeRegion('AzureInitialization');
  psm1.removeRegion('Finalization');

  if (project.azure) {
    psm1.setRegion('AzureInitialization', `
    # GS Testing
    $module = ipmo -passthru -ea 0 "C:\\work\\2018\\mark-powershell\\src\\Package\\Debug\\ResourceManager\\AzureResourceManager\\AzureRM.Profile.Netcore\\AzureRM.Profile.Netcore.psd1"

    # from PSModulePath
    # $module = ipmo -passthru -ea 0 "AzureRM.Profile.Netcore"

    Write-Host "Loaded Common Module '$($module.Name)'"

    # ask for the table of functions we can call in the common module.
    $VTable = Register-AzureModule

    # delegate responsibility to the common module for tweaking the pipeline at module load
    $instance.OnModuleLoad = $VTable.OnModuleLoad

    # and a chance to tweak the pipeline when we are about to make a call.
    $instance.OnNewRequest = $VTable.OnNewRequest

    # Need to get parameter values back from the common module
    $instance.GetParameterValue = $VTable.GetParameterValue

    # need to let the common module listen to events from this module
    $instance.EventListener = $VTable.EventListener
`);
  }

  psm1.setRegion('Initialization', `
    # this module instance.
    $instance =  [${project.serviceNamespace.moduleClass.declaration}]::Instance

    # load nested script module if it exists
    if( test-path "$PSScriptRoot/bin/${project.moduleName}.scripts.psm1" )  {
        ipmo "$PSScriptRoot/bin/${project.moduleName}.scripts.psm1"
    }

    $privatemodule = ipmo -passthru "$PSScriptRoot/bin/${project.moduleName}.private.dll"
    # export the 'exported' cmdlets
    Get-ChildItem "$PSScriptRoot/exported" -Recurse -Filter "*.ps1" -File | Sort-Object Name | Foreach {
        Write-Verbose "Dot sourcing private script file: $($_.Name)"
        . $_.FullName
        # Explicity export the member
        Export-ModuleMember -Function $_.BaseName
    }`);

  psm1.setRegion('Finalization', `
    # finish initialization of this module
    $instance.Init();
  `, false);

  psm1.trim();

  service.WriteFile(project.psm1, `${psm1}`, undefined, 'source-file-powershell');
}
