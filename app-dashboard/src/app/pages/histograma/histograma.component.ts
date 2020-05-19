import { Component, OnInit } from '@angular/core';
import { NorthwindService } from 'src/app/services/northwind.service';
import { Label } from 'ng2-charts';
import { ChartDataSets } from 'chart.js';
import { Observable } from 'rxjs/internal/Observable';

@Component({
  selector: 'app-histograma',
  templateUrl: './histograma.component.html',
  styleUrls: ['./histograma.component.scss']
})
export class HistogramaComponent implements OnInit {

  constructor(private north: NorthwindService) { }

  dataDimension: Label[] = [];
  dataValues: ChartDataSets[] = [];
  producto: Producto = new Producto();
  arrayProducto: Producto[]=[];

  //Data Ng Select
  defaultBindingsList = [
    { value: 1, label: 'Cliente', dimension: '[Dim Cliente].[Dim Cliente Nombre]' },
    { value: 2, label: 'Producto', dimension: '[Dim Producto].[Dim Producto Nombre]' },
    { value: 3, label: 'Empleado', dimension: '[Dim Empleado].[Dim Empleado Nombre]' }
  ];
  /*defaultBindingsAnnio = [
    { value: 1, label: '1996', dimension: '[Dim Tiempo].[Dim Tiempo Año].[1996]' },
    { value: 2, label: '1997', dimension: '[Dim Tiempo].[Dim Tiempo Año].[1997]' },
    { value: 3, label: '1998', dimension: '[Dim Tiempo].[Dim Tiempo Año].[1998]' }
  ];*/

  selectedDimension = null;
  customer$: Observable<any>;
  years$: Observable<any>;
  meses$: Observable<any>;
  selectedCustomer: string[] = [];
  selectedAnnio: string[] = [];
  selectedMes: string[] = [];
  ngOnInit(): void {

    this.selectedDimension = this.defaultBindingsList[0];

    this.north.getSerieHistorica(this.selectedDimension.label).subscribe((result: any)=>{
      //console.log(result.datosTabla);
      this.dataDimension = result.datosDimension;
      this.dataValues = result.datosTabla;
    });
    this.dataDimension = [];
    this.dataValues = [];

    this.customer$ = this.north.getItemsByDimension(`${this.selectedDimension.dimension}`, 'ASC');
    this.years$ = this.north.getItemsByDimension('[Dim Tiempo].[Dim Tiempo Año]', 'ASC');
    this.meses$ = this.north.getItemsByDimension('[Dim Tiempo].[Dim Tiempo Mes]', 'DESC');
  }

  recibeEvento($event){
    console.log($event);
  }

  onChangeDimension($event){
    //Cuestiones de buggeo
    //console.log('Item Seleccionado: ', event);
    this.selectedDimension = $event;
    
    /*this.north.getSerieHistorica(this.selectedDimension.label).subscribe((result: any)=>{
      //console.log(result.datosTabla);
      this.dataDimension = result.datosDimension;
      this.dataValues = result.datosTabla;
    });*/
    this.producto = new Producto();
    this.arrayProducto = [];
    this.producto.Mese = this.selectedMes;
    this.producto.Annio = this.selectedAnnio;
    this.producto.Nombre = this.selectedCustomer;
    this.north.getSerieHDimension(this.selectedDimension.dimension, 'DESC',this.arrayProducto).subscribe((result: any)=> {
      console.log(result.datosTabla);
      this.dataDimension = result.datosDimension;
      this.dataValues = result.datosTabla;
    });

    this.customer$ = this.north.getItemsByDimension(`${this.selectedDimension.dimension}`, 'ASC');
    this.years$ = this.north.getItemsByDimension('[Dim Tiempo].[Dim Tiempo Año]', 'ASC');
    this.meses$ = this.north.getItemsByDimension('[Dim Tiempo].[Dim Tiempo Mes]', 'DESC');
  }

  onChangeValues() {
    this.producto = new Producto();
    this.arrayProducto = [];
    this.producto.Mese = this.selectedMes;
    this.producto.Annio = this.selectedAnnio;
    this.producto.Nombre = this.selectedCustomer;
    console.log(this.producto.Mese);
    this.arrayProducto.push(this.producto);
    console.log(this.arrayProducto);
    this.north.getSerieHDimension(this.selectedDimension.dimension, 'DESC', this.arrayProducto).subscribe((result: any)=> {
      console.log(result)
      this.dataDimension = result.datosDimension;
      this.dataValues = result.datosTabla;
    });
  }

  clearModel() {
    this.selectedCustomer = [];
    this.selectedAnnio = [];
    this.selectedMes = [];
  }

  onChangeAnnio() {
    this.producto = new Producto();
    this.arrayProducto = [];
    this.producto.Mese = this.selectedMes;
    this.producto.Annio = this.selectedAnnio;
    this.producto.Nombre = this.selectedCustomer;
    console.log(this.producto.Mese);
    this.arrayProducto.push(this.producto);
    console.log(this.arrayProducto);
    this.north.getSerieHDimension(this.selectedDimension.dimension, 'DESC', this.arrayProducto).subscribe((result: any)=> {
      console.log(result)
      this.dataDimension = result.datosDimension;
      this.dataValues = result.datosTabla;
    });
  }
  onChangeMes() {
    this.producto = new Producto();
    this.arrayProducto = [];
    this.producto.Mese = this.selectedMes;
    this.producto.Annio = this.selectedAnnio;
    this.producto.Nombre = this.selectedCustomer;
    console.log(this.producto.Mese);
    this.arrayProducto.push(this.producto);
    console.log(this.arrayProducto);
    this.north.getSerieHDimension(this.selectedDimension.dimension, 'DESC', this.arrayProducto).subscribe((result: any)=> {
      console.log(result)
      this.dataDimension = result.datosDimension;
      this.dataValues = result.datosTabla;
    });
  }
}
export class Producto{
  Nombre: string[]=[];
  Annio: string[]=[];
  Mese: string[]=[];
}