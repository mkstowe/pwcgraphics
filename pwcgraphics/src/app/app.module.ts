import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { TopnavComponent } from './topnav/topnav.component';
import { ShopPageComponent } from './shop-page/shop-page.component';
import { ContactPageComponent } from './contact-page/contact-page.component';
import { ResourcesComponent } from './resources/resources.component';
import { HomePageComponent } from './home-page/home-page.component';
import { SocialsComponent } from './socials/socials.component';
import { SidenavComponent } from './sidenav/sidenav.component';

@NgModule({
  declarations: [
    AppComponent,
    TopnavComponent,
    ShopPageComponent,
    ContactPageComponent,
    ResourcesComponent,
    HomePageComponent,
    SocialsComponent,
    SidenavComponent,
  ],
  imports: [
    BrowserModule,
    AppRoutingModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
