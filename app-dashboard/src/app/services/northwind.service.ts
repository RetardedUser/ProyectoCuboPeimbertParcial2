import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { environment } from 'src/environments/environment';
import { Observable } from 'rxjs';
import { map } from 'rxjs/operators';
import { Producto } from '../pages/histograma/histograma.component';

const URL_API = environment.API.EndPoint.Northwind;

@Injectable({
  providedIn: 'root'
})
export class NorthwindService {

  constructor(private http: HttpClient) { 

  }

  getTop5(dimension: string){
    //this.http.get(URL_API+'Top5/'+dimension+'/');
    //modo o llamado asincrono
    /*return this.http.get(`${URL_API}Top5/${dimension}/`).subscribe((result:any)=>{

    });*/
    return this.http.get(`${URL_API}Top5/${dimension}/`);
  }

  getSerieHistorica(dimension: string){

    return this.http.get(`${URL_API}SerieH/${dimension}/`);
  }

  /*getCustomers(){
    return this.http.get(`./assets/JSON/clientes.json`);
  }*/
  getItemsByDimension(dimension: string, orden: string){
    return this.http.get(`${URL_API}GetItemsByDimension/${dimension}/${orden}`).pipe(map((result:any)=> result.datosDimension));
  }

  getDataPieByDimension(dimension: string, orden: string, values: string[]) {

    console.log('VALUES: ', values);

    return this.http.post(`${URL_API}GetDataPieByDimension/${dimension}/${orden}`, values).pipe(
      map((result: any) => result)
    )
  }
  
  getSerieHDimension(dimension: string, orden: string, values: Producto[]) {

    return this.http.post(`${URL_API}SerieHDimension/${dimension}/[Dim Tiempo].[Dim Tiempo Año]/[Dim Tiempo].[Dim Tiempo Mes]/${orden}`, values).pipe(
      map((result: any[]) => result)
    )
  }

  getPieDimension(dimension: string, orden: string, values: Producto[]) {

    return this.http.post(`${URL_API}PieDimension/${dimension}/[Dim Tiempo].[Dim Tiempo Año]/[Dim Tiempo].[Dim Tiempo Mes]/${orden}`, values).pipe(
      map((result: any[]) => result)
    )
  }
}
