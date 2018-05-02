import { TypeDeclaration } from './type-declaration';

export class LibraryType implements TypeDeclaration {
  constructor(private fullName: string) {
  }

  public get use(): string {
    return `${this.fullName}`;
  }

  public get implementation(): string {
    return ``;
  }
}

export const Unknown: TypeDeclaration = new LibraryType("null");
export const ToDo: TypeDeclaration = new LibraryType("null");
export const Void: TypeDeclaration = new LibraryType("void");
export const String: TypeDeclaration = new LibraryType("string");
export const Int: TypeDeclaration = new LibraryType("int");
export const Long: TypeDeclaration = new LibraryType("long");
export const Double: TypeDeclaration = new LibraryType("double");
export const Float: TypeDeclaration = new LibraryType("float");
export const Date: TypeDeclaration = new LibraryType("DateTime");
export const Duration: TypeDeclaration = new LibraryType("TimeSpan");
export const Binary: TypeDeclaration = new LibraryType("byte[]");
export const Bool: TypeDeclaration = new LibraryType("bool");
export const Object: TypeDeclaration = new LibraryType("object");
export const ThisObject: TypeDeclaration = new LibraryType("this object");

export const Task: TypeDeclaration = new LibraryType("System.Threading.Tasks.Task");
export const CancellationToken: TypeDeclaration = new LibraryType("System.Threading.CancellationToken");
export const HttpRequestMessage: TypeDeclaration = new LibraryType("System.Net.Http.HttpRequestMessage");
