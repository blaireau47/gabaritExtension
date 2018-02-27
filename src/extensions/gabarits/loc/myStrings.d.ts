declare interface IGabaritsCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'GabaritsCommandSetStrings' {
  const strings: IGabaritsCommandSetStrings;
  export = strings;
}
