import { CommaChar } from '#common/text-manipulation';
import { Class } from './class';
import { Method } from './method';

export class Constructor extends Method {
  constructor(protected containingClass: Class, objectIntializer?: Partial<Method>) {
    super(containingClass.name);
    this.apply(objectIntializer);
  }

  public get declaration(): string {
    const parameterDeclaration = this.parameters.joinWith(p => p.declaration, CommaChar);

    return `
${this.summaryDocumentation}
${this.parameterDocumentation}
${this.access} ${this.static} ${this.abstract} ${this.name}(${parameterDeclaration})
`.slim();
  }
}