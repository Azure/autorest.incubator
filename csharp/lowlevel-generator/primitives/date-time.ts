import { TypeDeclaration } from "#csharp/code-dom/type-declaration";

export class DateTime implements TypeDeclaration {
  get implementation(): string {
    return `System.DateTime`;
  };
  get use(): string {
    return `System.DateTime`
  }
  validation(propertyName: string): string {
    throw new Error("Method not implemented.");
  }
}

export class DateTime1123 implements TypeDeclaration {
  get implementation(): string {
    return `System.DateTime`;
  };
  get use(): string {
    return `System.DateTime`
  }
  validation(propertyName: string): string {
    throw new Error("Method not implemented.");
  }
}
