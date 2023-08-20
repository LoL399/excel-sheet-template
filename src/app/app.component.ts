import { Component } from '@angular/core';
import { toOTTemplate } from './OTTemplate';
import { toInOutTemplate } from './INOutTemplate';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'sheet-learn';

  async toExcel() {
    await toInOutTemplate();
  }
}
