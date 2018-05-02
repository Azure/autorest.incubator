import { EOL, comment, docCommentPrefix, indent, sortByName } from '#common/text-manipulation';
import { Namespace } from './namespace';
import { Type } from './type';

export class Interface extends Type {
  constructor(parent: Namespace, name: string, objectIntializer?: Partial<Interface>) {
    super(parent, name);
    this.apply(objectIntializer);
    parent.addInterface(this);
  }

  public get implementation(): string {
    const colon = this.interfaces.length > 0 ? ' : ' : '';
    const implementsInterfaces = this.interfaces.map(v => v.fullName).join(', ');
    const description = comment(this.description, docCommentPrefix);
    const methods = this.methods.sort(sortByName).map(m => m.interfaceDeclaration).join(EOL);
    const properties = this.properties.sort(sortByName).map(m => m.declaration).join(EOL);

    return `
${description}
${this.accessModifier} interface ${this.name}${colon}${implementsInterfaces} {
${indent(properties, 1)}
${indent(methods, 1)}
}
`.trim();
  }

  public get use(): string {
    return this.fullName;
  }

}
