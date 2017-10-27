/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.iesvdc.acceso.excelapi.excelapi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * Esta clase almacena información de libros para generar ficheros de Excel.
 * Un libro se compone de hojas
 * @author Carlos Aguilar Ortega
 */
public class Libro {
    /**
     * Atributos de la clase
     */
    private List <Hoja> hojas;
    private String nombreArchivo;
    
    
   /**
    * Constructor pasado un nombre
    * @param nombreArchivo 
    */
    public Libro(String nombreArchivo) {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = nombreArchivo;
    }
    /**
     * Constructor por defecto, ya que no s ele pasan parametros
     */
    public Libro(){
        this.hojas = new ArrayList<>();
        this.nombreArchivo = "nuevo.xlsx";
    }
    /**
     * Constructor pasandole el array de hojas y el nombre
     * @param hojas
     * @param nombreArchivo 
     */
    public Libro(ArrayList<Hoja> hojas, String nombreArchivo) {
        this.hojas = hojas;
        this.nombreArchivo = nombreArchivo;
    }
    
    public List<Hoja> getHojas() {
        return hojas;
    }

    public void setHojas(List<Hoja> hojas) {
        this.hojas = hojas;
    }

    public String getNombreArchivo() {
        return nombreArchivo;
    }

    public void setNombreArchivo(String nombreArchivo) {
        this.nombreArchivo = nombreArchivo;
    }
    
    public boolean addHoja(Hoja hoja){
        return this.hojas.add(hoja);
    }
    
    public Hoja removeHoja(int index) throws ExcelAPIException {
        if(index < 0 || index>this.hojas.size()){
            throw new ExcelAPIException("Libro::removeHoja(): Posicion no válida");
        }
       return this.hojas.remove(index);
    }
    
    public Hoja indexHoja(int index) throws ExcelAPIException {
        if(index < 0 || index>this.hojas.size()){
            throw new ExcelAPIException("Libro::indexHoja(): Posicion no válida");
        }
        return this.hojas.get(index);
    }
    
    public void load(){
        
    }
    
    public void load(String fileName){
        this.nombreArchivo = fileName;
        this.load();
    }
    
    public void save() throws ExcelAPIException{
        SXSSFWorkbook wb = new SXSSFWorkbook();
        for (Hoja hoja : this.hojas) {
            Sheet sh = wb.createSheet(hoja.getNombre());
            for (int i = 0; i < hoja.getnFilas(); i++) {
                Row row = sh.createRow(i);
                for (int j = 0; j < hoja.getnColumnas(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(hoja.getDato(i, j));                
                }
            }
        }
        try {
            FileOutputStream out = new FileOutputStream(this.nombreArchivo);
            wb.write(out);
            out.close();                        
        } catch (IOException ex) {
           throw new ExcelAPIException("Error al guardar el fichero");
        } finally {
            wb.dispose();
        }
    }
    
    public void save(String fileName) throws ExcelAPIException{
        this.nombreArchivo = fileName; 
        this.save();
    }
    
    private void testExtension(){
        if(!this.nombreArchivo.matches("^(?:[\\w]\\:|\\\\)(\\[a-z_\\-\\s0-9\\.]+)+\\.(txt|gif|pdf|doc|docx|xls|xlsx)$")){
            this.nombreArchivo=this.nombreArchivo+".xlsx";
        }
    }
}