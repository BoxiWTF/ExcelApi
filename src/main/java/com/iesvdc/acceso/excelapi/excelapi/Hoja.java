/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.iesvdc.acceso.excelapi.excelapi;

/**
 *  Esta clase almacena información del texto de una hoja de cálculo
 * 
 * @author Carlos Aguilar Ortega
 */
public class Hoja {
    private String [][] datos;
    private String nombre;
    private int nFilas;
    private int nColumnas;
    /**
     * Crea una hoja de calculo nueva
     */
    public Hoja() {
        this.datos = new String[5][5];
        this.nFilas=5;
        this.nColumnas=5;
        this.nombre = "";
    }
    /**
     * Crea una hoja de tamaño nFilas por nColumnas
     * @param nFilas numero de filas
     * @param nColumnas numero de columnas
     */
    public Hoja(int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        nombre="";
        this.nFilas=nFilas;
        this.nColumnas=nColumnas;
    }
    /**
     * Crea una hoja de tamano nFilas y nColumnas y ademas se le añade un nombre
     * @param nombre el nombre de la hoja
     * @param nFilas el numero de filas
     * @param nColumnas el numero de columnas
     */
    public Hoja(String nombre, int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        this.nombre = nombre;
        this.nFilas=nFilas;
        this.nColumnas=nColumnas;
    }

    public String getDato(int fila, int columna) {
        return datos[fila][columna];
        
    }

    public void setDato(String dato, int fila, int columna) {
        /*TO-DO manejar si accedemos a una posicion no valida*/
        this.datos[fila][columna] = dato;
    }
    
    public String getNombre(){
        return this.nombre;
    }
    
    public void setNombre(String nombre){
        this.nombre = nombre;
    }

    public int getnFilas() {
        return nFilas;
    }

    public int getnColumnas() {
        return nColumnas;
    }
    
    public boolean compare(Hoja hoja){
        boolean iguales = true;
        if(this.nColumnas==hoja.nColumnas && this.nFilas==hoja.nFilas && this.nombre==hoja.getNombre()){
            for(int i=0; i<this.nFilas; i++){
                for(int j=0; j<this.nColumnas; j++){
                    if(!this.datos[i][j].equals(hoja.getDato(i, j))){
                        return false;
                    }
                }
            }
        }else{
            iguales = false;
        }
        return iguales;
    }
}
