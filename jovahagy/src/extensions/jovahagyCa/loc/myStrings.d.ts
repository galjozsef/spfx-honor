declare interface IJovahagyCaCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'JovahagyCaCommandSetStrings' {
  const strings: IJovahagyCaCommandSetStrings;
  export = strings;
}
