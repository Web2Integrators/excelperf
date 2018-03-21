import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

declare const Office: any;

Office.initialize = (reason) => {
  console.log(reason);
platformBrowserDynamic().bootstrapModule(AppModule)
    .catch(err => console.log(err));
};

// platformBrowserDynamic().bootstrapModule(AppModule)
//   .catch(err => console.log(err));
