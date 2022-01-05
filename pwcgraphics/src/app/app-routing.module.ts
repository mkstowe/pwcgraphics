import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { ContactPageComponent } from './contact-page/contact-page.component';
import { HomePageComponent } from './home-page/home-page.component';
import { ResourcesComponent } from './resources/resources.component';
import { ShopPageComponent } from './shop-page/shop-page.component';

const routes: Routes = [
  { path: '', component: HomePageComponent, data: { tab: 1 } },
  { path: 'shop', component: ShopPageComponent, data: { tab: 2 } },
  { path: 'contact', component: ContactPageComponent, data: { tab: 4 } },
  { path: 'resources', component: ResourcesComponent, data: { tab: 5 } },
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
