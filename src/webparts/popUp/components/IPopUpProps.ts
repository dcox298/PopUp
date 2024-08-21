export interface IPopUpProps {
  buttonText:string;
  popUpText:string;
  isDarkTheme: boolean;

  buttonType:string;
  buttonAlignment:"auto" | "center" | "baseline" | "stretch" | "start" | "end"|undefined;

  onTextChange?:any;
  displayMode:number

  backgroundColor:string;
}
