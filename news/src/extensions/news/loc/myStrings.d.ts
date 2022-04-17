declare interface INewsCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module "NewsCommandSetStrings" {
  const strings: INewsCommandSetStrings;
  export = strings;
}
