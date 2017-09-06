import { Component } from '@angular/core';;
declare const Excel: any;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  menus:{}[] = [
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  {name:'first', icon:'fa-home'},
  ];

  firstMacro1(){
  	Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = 'blue';
      // this.ListAllSheetNames(context);
      await context.sync();
    });
  }

  firstMacro(ctx) {
    Excel.run(async (ctx) => {
      // let sheet = ctx.workbook.worksheets.getActiveWorksheet();
      // let values = [[]];
      //ctx.workbook.worksheets
      // const range = context.workbook.getSelectedRange();
      // range.format.fill.color = 'blue';
      // const AR = context.workbook.getCell().row;
      // const AC = context.workbook.getCell().column;
      // const AS = context.workbook.getActiveWorksheet();
      // const AS1 = context.workbook.
      // for (let sheet of ctx.workbook.worksheets) {
      //   values[0].push(sheet.name);
      // }
      // values[0].push(1);
      // let range = sheet.getRange("A1:A1");
      // range.values = values;

      // let values = [["Product"]];
      //
      // //Queue a command to write the sample data to the specified range
      // //in the worksheet and bold the header row
      // let range = sheet.getRange("A1");
      // range.values = values;

      const range = ctx.workbook.getSelectedRange();
      let v = [[]];
      for (let i = 0; i < 1000; i++){
        v.push([]);
        for (let j = 0; j < 1000; j++){
          v[i].push(j);
        }
      }
      range.values = v;

      await ctx.sync();
    });
  }

}
