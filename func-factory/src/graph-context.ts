import { SymbolSource, Graph, NodePhi, NodeProc, ControlSource, ControlSink, SymbolSink, ControlFlow, DataFlow } from './graph';
import {
  getNodes,
  getControlSources,
  getControlSinks,
  getSymbolSinks,
  getSymbolSources,
  getSymbolSourceOf,
  getSymbolMapper,
  getMarket,
  nodeSource2Sinks,
  nodeSource2Sources,
  nodeSink2Sinks,
  nodeSink2Sources,
  getSymbolSinkType,
  getSymbolSourceType,
  getSymbolSinkKnownNames
} from './graph-analysis';
import { validateNodeProc, validateNodePhi, validateRawControlFlow, validateControlFlow, validateRawDataFlow, validateDataFlow, validateSymbolAvailability } from './graph-validation';
import { objMap, setExcept, lundef, clone, objReplace, tsc, Obj, trycatch, error, deepEquals, errorUnreachable } from './helpers';
import { ProcImplementation, generateTS, GenerationFlavor, ProcImplementations } from './reference-generator';

export interface Proc<TType> {
  readonly pure: boolean, // side effect free?
  readonly inputs: { readonly [id: string]: { readonly names: ReadonlyArray<string>, readonly type: TType } | undefined }; // data sink
  readonly outputFlows: { readonly [flow: string]: { readonly [id: string]: { readonly names: ReadonlyArray<string>, readonly nameSources: ReadonlyArray<string>, readonly type: TType } | undefined } | undefined }; // data source
}
export interface ProcDefinitions<TType> {
  readonly [id: string]: Proc<TType> | undefined;
}

export interface SymbolInstance<TType> {
  readonly source: SymbolSource<TType>;
  readonly type: TType;
  readonly names: ReadonlyArray<string>;
}

export type GraphProblem = {
  readonly severity: "error" /* would generate invalid code */ | "warning" /* can generate working code */;
  readonly needsHumanIntervention: boolean; // cannot (theoretically) be resolved by graph synthesis
  readonly message: string;
  readonly graphComponent: any;
}

export type FlexArgs = Obj<any>;
export type FlexCallbacks = Obj<(args: FlexArgs) => void>;
export type FlexFunc = (args: FlexArgs, cbs: FlexCallbacks) => void;
export type FlexCallbacksAsync = Obj<(args: FlexArgs) => Promise<void>>;
export type FlexFuncAsync = (args: FlexArgs, cbs: FlexCallbacks) => Promise<void>;

export interface Sample {
  input: FlexArgs;
  output: FlexArgs;
  outputFlow: string;
  impl: ProcImplementations;
}

/**
 * Operations for validating, analyzing and manipulating graphs.
 */
export class GraphContext<TType> {
  // vertices
  public readonly nodesPhi: ReadonlyArray<NodePhi<TType>>;
  public readonly nodesProc: ReadonlyArray<NodeProc<TType>>;
  // control flow connectors
  public readonly controlSources: ReadonlyArray<ControlSource<TType>>;
  public readonly controlSinks: ReadonlyArray<ControlSink<TType>>;
  private readonly controlSourceNorm: (x: ControlSource<TType>) => ControlSource<TType> | undefined;
  private readonly controlSinkNorm: (x: ControlSink<TType>) => ControlSink<TType> | undefined;
  // data flow connectors
  public readonly symbolSources: ReadonlyArray<SymbolSource<TType>>;
  public readonly symbolSinks: ReadonlyArray<SymbolSink<TType>>;
  private readonly symbolSourceNorm: (x: SymbolSource<TType>) => SymbolSource<TType> | undefined;
  private readonly symbolSinkNorm: (x: SymbolSink<TType>) => SymbolSink<TType> | undefined;
  // control flow
  public readonly controlFlow: ReadonlyArray<ControlFlow<TType>>;
  // data flow
  public readonly dataFlow: ReadonlyArray<DataFlow<TType>>;
  // symbols
  public symbols: ReadonlyArray<SymbolInstance<TType>>;
  public getSymbolFromSource: (src: SymbolSource<TType>) => SymbolInstance<TType> | undefined;
  private supply: ReadonlyMap<ControlSource<TType>, ReadonlySet<SymbolInstance<TType>>>;
  private demand: ReadonlyMap<ControlSink<TType>, ReadonlySet<SymbolInstance<TType>>>;
  public getSupply(x: ControlSource<TType>): ReadonlySet<SymbolInstance<TType>> { return lundef(this.controlSourceNorm(x), x => this.supply.get(x)) || new Set<SymbolInstance<TType>>(); }
  public getDemand(x: ControlSink<TType>): ReadonlySet<SymbolInstance<TType>> { return lundef(this.controlSinkNorm(x), x => this.demand.get(x)) || new Set<SymbolInstance<TType>>(); }

  // opportunities
  public unconnectedControlSources: ReadonlyArray<ControlSource<TType>>;
  public unconnectedControlSinks: ReadonlyArray<ControlSink<TType>>;
  public unconnectedSymbolSources: ReadonlyArray<SymbolSource<TType>>;
  public unconnectedSymbolSinks: ReadonlyArray<SymbolSink<TType>>;

  // quality
  public readonly problems: ReadonlyArray<GraphProblem>;
  public get canGenerateWorkingCode(): boolean {
    return this.problems.every(p => p.severity !== "error");
  }
  public get canBeFixedWithSynthesis(): boolean {
    return this.problems.every(p => !p.needsHumanIntervention);
  }

  private _score: number | undefined = undefined;
  public get score(): number {
    return this._score || (this._score = (() => {
      if (!this.canBeFixedWithSynthesis) return 0;
      if (this.canGenerateWorkingCode) return this.matchesSamples() ? 1 : 0;
      return Math.pow(0.75, this.unconnectedControlSources.length)
        * Math.pow(0.9, this.unconnectedControlSinks.length)
        * (1 - 0.1 * Math.pow(0.9, this.dataFlow.map(x => this.nameProximity(getSymbolSinkKnownNames(this.procs, x.target), lundef(this.getSymbolFromSource(x.source), _ => _.names) || [])).reduce((a, b) => a + b, 0)))
        * Math.pow(0.9, Math.abs(this.unconnectedSymbolSources.length - this.unconnectedSymbolSinks.length))
        * Math.pow(0.95, this.unconnectedSymbolSources.length + this.unconnectedSymbolSinks.length)
        * Math.pow(0.95, this.problems.length);
    })());
  }

  public matchesSamples(onMessage: (message: string, sample: Sample | null) => void = () => { }): boolean {
    if (!this.canGenerateWorkingCode) {
      onMessage("Cannot generate code", null);
      return false;
    }
    let result = true;
    for (const sample of this.samples) {
      const func = this.build(sample.impl);

      const cbs: FlexCallbacks = objMap(fl => (args: any) => {
        onMessage("Unexpected output flow: " + fl, sample);
        result = false;
      }, this.graph.outputFlows);
      cbs[sample.outputFlow] = (args: any) => {
        if (!deepEquals(args, sample.output)) {
          onMessage(`Output mismatch: ${JSON.stringify(args)} vs ${sample.output}`, sample);
          result = false;
        }
      };
      func(sample.input, cbs);
    }
    return result;
  }

  private updateGraph(graph: Graph<TType>): GraphContext<TType> {
    return new GraphContext(graph, this.typeAssignableTo, this.typeToTS, this.procs, this.samples);
  }

  public constructor(
    public readonly graph: Graph<TType>,
    public readonly typeAssignableTo: (from: TType, to: TType) => boolean,
    public readonly typeToTS: (type: TType) => string,
    public readonly procs: ProcDefinitions<TType>,
    public readonly samples: Sample[] = []) {

    const problems: GraphProblem[] = [];
    this.problems = problems;
    const onProblem = (problem: GraphProblem): void => { problems.push(problem) };

    // vertices
    const { nodesPhi, nodesProc } = getNodes(graph);
    this.nodesPhi = [...nodesPhi];
    this.nodesProc = [...nodesProc];
    // (validate)
    for (const nodePhi of nodesPhi) validateNodePhi(nodePhi, onProblem);
    for (const nodeProc of nodesProc) validateNodeProc(nodeProc, procs, onProblem);

    // control flow connectors
    const { controlSources, controlSourceEquals } = getControlSources(procs, nodesPhi, nodesProc);
    this.controlSources = controlSources;
    this.controlSourceNorm = x => controlSources.find(y => controlSourceEquals(x, y));
    const { controlSinks, controlSinkEquals } = getControlSinks(graph.outputFlows, procs, nodesPhi, nodesProc);
    this.controlSinks = controlSinks;
    this.controlSinkNorm = x => controlSinks.find(y => controlSinkEquals(x, y));

    // data flow connectors
    const { symbolSources, symbolSourceEquals } = getSymbolSources(graph.inputs, procs, controlSources, controlSourceEquals);
    this.symbolSources = symbolSources;
    this.symbolSourceNorm = x => symbolSources.find(y => symbolSourceEquals(x, y));
    const { symbolSinks, symbolSinkEquals } = getSymbolSinks(graph.outputFlows, procs, controlSinks, controlSinkEquals);
    this.symbolSinks = symbolSinks;
    this.symbolSinkNorm = x => symbolSinks.find(y => symbolSinkEquals(x, y));

    // control flow edges
    this.controlFlow = graph.controlFlow.map(x => ({ source: this.controlSourceNorm(x.source), target: this.controlSinkNorm(x.target) }))
      .filter((x): x is ControlFlow<TType> => x.source !== undefined && x.target !== undefined);
    // (validate)
    validateRawControlFlow(graph.controlFlow, this.controlSourceNorm, this.controlSinkNorm, onProblem);
    validateControlFlow(this.controlFlow, controlSources, controlSinks, onProblem);

    // data flow edges
    this.dataFlow = graph.dataFlow.map(x => ({ source: this.symbolSourceNorm(x.source), target: this.symbolSinkNorm(x.target) }))
      .filter((x): x is DataFlow<TType> => x.source !== undefined && x.target !== undefined);
    // (validate)
    validateRawDataFlow(graph.dataFlow, procs, symbolSinks, this.symbolSourceNorm, this.symbolSinkNorm, onProblem);
    validateDataFlow(graph, procs, this.dataFlow, this.symbolSources, this.symbolSinks, this.typeAssignableTo, onProblem);

    // symbols
    this.getSymbolFromSource = getSymbolMapper(graph.inputs, this.dataFlow, procs, symbolSources, this.symbolSourceNorm, this.symbolSinkNorm);
    this.symbols = symbolSources.map(this.getSymbolFromSource).filter((x): x is SymbolInstance<TType> => x !== undefined);
    const { demand, supply } = getMarket(this.dataFlow, this.controlFlow, controlSources, controlSinks, symbolSources, this.controlSourceNorm, this.controlSinkNorm, this.getSymbolFromSource);
    this.supply = supply;
    this.demand = demand;
    // (validate that symbols are available where required)
    validateSymbolAvailability(this.controlFlow, supply, demand, onProblem);

    // opportunities
    this.unconnectedControlSources = [...setExcept(this.controlSources, this.controlFlow.map(x => x.source))];
    this.unconnectedControlSinks = [...setExcept(this.controlSinks, this.controlFlow.map(x => x.target))];
    this.unconnectedSymbolSources = [...setExcept(this.symbolSources, this.dataFlow.map(x => x.source))];
    this.unconnectedSymbolSinks = [...setExcept(this.symbolSinks, this.dataFlow.map(x => x.target))];
  }

  // traversal
  public nodeSource2Sinks(x: ControlSource<TType>): ReadonlyArray<ControlSink<TType>> {
    return nodeSource2Sinks(x, this.controlSinks);
  }

  public nodeSource2Sources(x: ControlSource<TType>): ReadonlyArray<ControlSource<TType>> {
    return nodeSource2Sources(x, this.controlSources);
  }

  public nodeSink2Sinks(x: ControlSink<TType>): ReadonlyArray<ControlSink<TType>> {
    return nodeSink2Sinks(x, this.controlSinks);
  }

  public nodeSink2Sources(x: ControlSink<TType>): ReadonlyArray<ControlSource<TType>> {
    return nodeSink2Sources(x, this.controlSources);
  }

  public edgeSink2Source(x: ControlSink<TType>): ControlSource<TType> | undefined {
    const y = this.controlSinkNorm(x);
    return this.controlFlow.filter(f => f.target === y).map(x => x.source)[0];
  }

  public edgeSource2Sink(x: ControlSource<TType>): ControlSink<TType> | undefined {
    const y = this.controlSourceNorm(x);
    return this.controlFlow.filter(f => f.source === y).map(x => x.target)[0];
  }

  public edgeSymbolSink2Source(x: SymbolSink<TType>): SymbolSource<TType> | undefined {
    const y = this.symbolSinkNorm(x);
    return this.dataFlow.filter(f => f.target === y).map(x => x.source)[0];
  }

  public edgeSymbolSource2Sink(x: SymbolSource<TType>): SymbolSink<TType> | undefined {
    const y = this.symbolSourceNorm(x);
    return this.dataFlow.filter(f => f.source === y).map(x => x.target)[0];
  }

  public getSymbolSinkType(x: SymbolSink<TType>): TType | undefined {
    return getSymbolSinkType(this.graph.outputFlows, this.procs, x);
  }

  public getSymbolSourceType(x: SymbolSource<TType>): TType | undefined {
    return getSymbolSourceType(this.graph.inputs, this.procs, x);
  }

  public getGraphProc(): Proc<TType> {
    return {
      pure: [...this.nodesProc].map(x => this.procs[x.procID]).every(x => x === undefined ? false : x.pure),
      inputs: objMap((_, value) => value && !value.value && ({ type: value.type, names: value.names }) || undefined, this.graph.inputs),
      outputFlows: objMap((_, ofTarget) => ofTarget && objMap((_, sy) => {
        if (sy === undefined) return undefined;
        return { names: [] /* TODO: could grab name candidates from data source */, type: sy, nameSources: [] };
      }, ofTarget), this.graph.outputFlows)
    };
  }

  // EXECUTION
  public compile(impls: ProcImplementations): string {
    return generateTS(this, impls, this.typeToTS, GenerationFlavor.ContInlineProc);
  }
  public build(impls: ProcImplementations): FlexFunc {
    const funcTs = this.compile(impls);
    const funcJs = trycatch(() => tsc(funcTs), ex => { throw new Error(ex + "\n\n" + funcTs); });
    return eval(funcJs);
  }

  // MUTATION
  public removeControlFlow(edge: ControlFlow<TType>): GraphContext<TType> {
    return this.updateGraph({
      inputs: this.graph.inputs,
      controlFlow: this.graph.controlFlow.filter(x => edge.source !== this.controlSourceNorm(x.source) || edge.target !== this.controlSinkNorm(x.target)),
      dataFlow: this.graph.dataFlow,
      outputFlows: this.graph.outputFlows
    });
  }
  public connectControlFlow(source: ControlSource<TType>, sink: ControlSink<TType>): GraphContext<TType> {
    return this.updateGraph({
      inputs: this.graph.inputs,
      controlFlow: this.graph.controlFlow.concat([{ source: source, target: sink }]),
      dataFlow: this.graph.dataFlow,
      outputFlows: this.graph.outputFlows
    });
  }
  public connectDataFlow(source: SymbolSource<TType>, sink: SymbolSink<TType>): GraphContext<TType> {
    return this.updateGraph({
      inputs: this.graph.inputs,
      controlFlow: this.graph.controlFlow,
      dataFlow: this.graph.dataFlow.concat([{ source: source, target: sink }]),
      outputFlows: this.graph.outputFlows
    });
  }

  /**
   * Calculates proximity score between two sets of names (associated with a symbol each).
   * Higher score means higher proximity and hence higher likelihood that the associated symbols are related - and could be plumbed together.
   * 
   * NOTE: This can be way more sophisticated, e.g. case insensitivity, partial score for substring match, etc. 
   */
  private nameProximity(names1: ReadonlyArray<string>, names2: ReadonlyArray<string>): number {
    let result = 0;
    for (const name of names1)
      if (names2.includes(name))
        result++;
    return result;
  }

  // SYNTHESIS
  public synthesizeNextGeneration(): Array<GraphContext<TType>> {
    // if (this.canGenerateWorkingCode) return [];
    if (this.unconnectedSymbolSinks.length === 0 && this.unconnectedControlSources.length === 0) return [];
    // TODO: grade based on name stuff

    // 1) connect ctrl flow
    if (this.unconnectedControlSources.length > 0) {
      // MAYBE: consider for ALL sinks? (inserting phi if necessary)
      return this.unconnectedControlSinks.map(sink => this.connectControlFlow(this.unconnectedControlSources[0], sink));
    }
    // 2) connect data flow 
    const symSink = this.unconnectedSymbolSinks.map(symSink => ({ symSink: symSink, ctrlSrc: this.edgeSink2Source(symSink.target) })).filter(x => x.ctrlSrc !== undefined)[0];
    const symSinkType = this.getSymbolSinkType(symSink.symSink) || errorUnreachable();
    const available = this.getSupply(symSink.ctrlSrc || errorUnreachable());
    const sym = [...available].filter(x => this.typeAssignableTo(x.type, symSinkType));
    const resultConnect = sym.map(x => this.connectDataFlow(x.source, symSink.symSink));
    // 3) insert operation 
    const edge = this.controlFlow.find(x => x.target === symSink.symSink.target) || errorUnreachable();
    const ga = this.removeControlFlow(edge);
    const resultAddProc = Object.keys(this.procs).map(procId => {
      const outFlows = lundef(this.procs[procId], _ => _.outputFlows) || {};
      return Object.keys(outFlows)
        .map(flowId => {
          const flow = outFlows[flowId] || {};
          const procNode: NodeProc<TType> = { procID: procId };
          return Object.entries(flow).filter(x => lundef(x[1], _ => this.typeAssignableTo(_.type, symSinkType)) || false)
            .map(x => x[0])
            .map(outId => ga
              .connectControlFlow(edge.source, { type: "proc", node: procNode })
              .connectControlFlow({ type: "proc", node: procNode, flow: flowId }, edge.target)
              .connectDataFlow({ origin: { type: "proc", node: procNode, flow: flowId }, id: outId }, symSink.symSink));
        })
        .reduce((a, b) => a.concat(b), []);
    }).reduce((a, b) => a.concat(b), []);

    return resultConnect.concat(resultAddProc);
  }

  public synthesize(maxPopulationSize: number | undefined = undefined): GraphContext<TType> | undefined {
    let population: Array<GraphContext<TType>> = [this];
    while (true) {
      const c = population.shift();
      if (c === undefined) return undefined;
      const nextGeneration = c.synthesizeNextGeneration();

      const solution = nextGeneration.find(x => x.score === 1);
      if (solution !== undefined) return solution;

      // // dump working code with imperfect score (helpful for seeing which direction the algorithm explores):
      // for (const x of nextGeneration.filter(x => x.canGenerateWorkingCode).map(x => x.compile(x.samples[0].impl)))
      //   console.log(x);
      population.push(...nextGeneration.filter(x => !x.canGenerateWorkingCode));

      const nsize = (ga: GraphContext<TType>) => ga.nodesPhi.length + ga.nodesProc.length;
      const nscore = (ga: GraphContext<TType>) => ga.score;
      population = population.sort((a, b) => nsize(a) == nsize(b) ? nscore(b) - nscore(a) : nsize(a) - nsize(b));
      if (maxPopulationSize !== undefined) population = population.slice(0, maxPopulationSize);
    }
  }
}
