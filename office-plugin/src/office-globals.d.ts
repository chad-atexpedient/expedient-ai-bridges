declare global {
  interface Window {
    Office?: any;
    Excel?: any;
    Word?: any;
    PowerPoint?: any;
  }

  var Office: any;
  var Excel: any;
  var Word: any;
  var PowerPoint: any;
}

export {};
