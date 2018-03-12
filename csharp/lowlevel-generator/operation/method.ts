import { Class } from "#csharp/code-dom/class";
import * as codemodel from "#remodeler/code-model";
import { Method } from "#csharp/code-dom/method";
import * as mscorlib from "#csharp/code-dom/mscorlib";
import { Parameter } from "#csharp/code-dom/parameter";
import { OperationParameter, OperationBodyParameter, CallbackParameter } from "../operation/parameter";

import { Model } from "#remodeler/code-model";
import * as message from "../messages";
import { all, EOL } from "#common/text-manipulation";
import { State } from "../generator";
import { Statement, Statements, OneOrMoreStatements } from "#csharp/code-dom/statements/statement";
import { Expression } from "#csharp/code-dom/expression";
import { Using } from "#csharp/code-dom/statements/using";
import { Try } from "#csharp/code-dom/statements/try";
import { Finally } from "#csharp/code-dom/statements/finally";
import { Switch } from "#csharp/code-dom/statements/switch";
import { Case, DefaultCase, TerminalDefaultCase } from "#csharp/code-dom/statements/case";

export class OperationMethod extends Method {
  private ready: Promise<OperationMethod>;
  protected constructor(name: string, protected parent: Class, protected operation: codemodel.HttpOperation, protected state: State) {
    super(name, mscorlib.Task);
    // finish in async constructor
    this.ready = (async () => {
      this.isAsync = true;

      // add parameters
      const methodParameters = await Promise.all(this.operation.parameters.map(async (value, index) => await OperationParameter.create(this, value, this.state.path('parameters', index))));

      this.parent.addMethod(this);
      this.description = this.operation.details.description || "";

      // add body paramter
      if (this.operation.requestBody) {
        const appjson = this.operation.requestBody.content["application/json"];
        if (appjson && appjson.schema) {
          await OperationBodyParameter.create(this, "body", this.operation.requestBody.description || "", appjson.schema, this.state.path('requestBody'));
        }
      }

      // add response delegate parameters
      for (const responseCode in this.operation.responses) {
        const response = this.operation.responses[responseCode];
        const rState = this.state.path('responses', responseCode);
        const c = response.content["application/json"];
        if (c) {
          await CallbackParameter.create(this, `on${responseCode}`, c.schema, this.state);
        } else {
          await CallbackParameter.create(this, `on${responseCode}`, undefined, this.state);
        }
      }
      // add cancellationToken parameter

      const cancellationToken = this.addParameter(new Parameter("cancellationToken", mscorlib.CancellationToken));

      // add eventhandler parameter
      const listener = this.addParameter(new Parameter("listener", mscorlib.EventListener));

      this.add(this.methodBody(methodParameters, listener, cancellationToken));
      return this;
    })();
  }

  public static async create(parent: Class, operation: codemodel.HttpOperation, state: State): Promise<OperationMethod> {
    return new OperationMethod(operation.details.name, parent, operation, state).ready;
  }

  private *methodBody(methodParameters: Array<OperationParameter>, listener: Parameter, cancellationToken: Parameter): Iterable<string | Statement> {
    const eventListener = new EventListener(listener);
    const operation = this.operation;

    yield `${listener.value}.cancellationToken = ${cancellationToken.value};`;

    yield `// fire event before validation`;
    yield eventListener.fire("Validation");

    yield EOL;
    yield `// perform parameter validation here`;

    for (const parameter of methodParameters) {
      // spit out parameter validation
      yield parameter.validate;
    }

    yield eventListener.fire("AfterValidation");


    yield new Using(listener.value, function* () {
      yield new Using(`var _request = new System.Net.Http.HttpRequestMessage(Microsoft.Rest.Method.${operation.method.capitalize()}, "http://wherever/...")`, function* () {
        yield eventListener.fire("RequestCreated", '_request');

        yield `System.Net.Http.HttpResponseMessage _response = null;`;
        yield new Try(function* () {
          // try statements
          yield `_response = await this.Sender.SendAsync(_request, ${cancellationToken.value}, ${listener.value});`;
          yield eventListener.fire("ResponseCreated", '_response');

          // add response handlers
          yield new Switch({ value: `_response.StatusCode` }, function* () {
            for (const responseCode in operation.responses) {
              if (responseCode !== 'default') {
                // each response
                yield new Case(`(System.Net.HttpStatusCode)${responseCode}`, function* () {
                  yield `// on ${responseCode} ... `;
                  yield EOL;
                });
              } else {
                yield new DefaultCase(function* () {
                  yield "// on default ... ";
                  yield EOL;
                });
              }
            }

            if (!operation.responses["default"]) {
              // if no default, we need one that handles the rest of the stuff.
              yield new TerminalDefaultCase(function* () {
                yield `throw new Microsoft.Rest.UndeclaredResponseException(_response.StatusCode);`;
              });
            }
          });
        });

        yield new Finally(function* () {
          yield "// finally statements";
          yield eventListener.fire("Finally", "_request", "_response");
          yield `_response?.Dispose();`
        });

      });

    });

  }
}

export class EventListener {
  constructor(protected expression: Expression) {
  }

  fire(eventName: string, ...additionalParameters: Array<string | Expression>) {
    return new FireEvent(this.expression, eventName, additionalParameters);
  }
}

export class FireEvent implements Statement {
  constructor(protected expression: Expression, protected eventName: string, protected additionalParameters: Array<string | Expression>) {
  }

  get implementation(): string {
    const additionalParameters = this.additionalParameters.length > 0 ? `, ${this.additionalParameters.joinWith(each => typeof each === 'string' ? each : each.value)}` : ``;

    return `await ${this.expression.value}.Signal(Microsoft.Rest.Events.${this.eventName}${additionalParameters});`;
  }
}
/*
export class FireEventWithRequest extends Statement {

}

export class HttpStatusCodeCase extends Case {

}*/