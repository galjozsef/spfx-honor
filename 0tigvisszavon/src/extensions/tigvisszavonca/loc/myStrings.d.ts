declare interface ITigvisszavoncaCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'TigvisszavoncaCommandSetStrings' {
  const strings: ITigvisszavoncaCommandSetStrings;
  export = strings;
}
