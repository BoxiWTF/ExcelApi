/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.iesvdc.acceso.excelapi.excelapi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import static org.apache.poi.ss.usermodel.CellType.BOOLEAN;
import static org.apache.poi.ss.usermodel.CellType.FORMULA;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;
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
    /**
     * Metodo que devuelve las hojas (en array)
     * @return 
     */
    public List<Hoja> getHojas() {
        return hojas;
    }
    /**
     * Metodo que setea un array de hojas
     * @param hojas 
     */
    public void setHojas(List<Hoja> hojas) {
        this.hojas = hojas;
    }
    /**
     * Metodo que devuelve el nombre del archivo
     * @return 
     */
    public String getNombreArchivo() {
        return nombreArchivo;
    }
    /**
     * Metodo que setea el nombre del archivo
     * @param nombreArchivo 
     */
    public void setNombreArchivo(String nombreArchivo) {
        this.nombreArchivo = nombreArchivo;
    }
    /**
     * Metodo para añadir una hoja
     * @param hoja
     * @return 
     */
    public boolean addHoja(Hoja hoja){
        return this.hojas.add(hoja);
    }
    /**
     * Metodo para borrar una hoja
     * @param index
     * @return
     * @throws ExcelAPIException 
     */
    public Hoja removeHoja(int index) throws ExcelAPIException {
        if(index < 0 || index>this.hojas.size()){
            throw new ExcelAPIException("Libro::removeHoja(): Posicion no válida");
        }
       return this.hojas.remove(index);
    }
    /**
     * Metodo que accede a una hoja del array de hojas
     * @param index
     * @return
     * @throws ExcelAPIException 
     */
    public Hoja indexHoja(int index) throws ExcelAPIException {
        if(index < 0 || index>this.hojas.size()){
            throw new ExcelAPIException("Libro::indexHoja(): Posicion no válida");
        }
        return this.hojas.get(index);
    }
    /**
     * Metodo load, cuya funcion es leer una hoja de apache POI y pasarla a una hoja.java
     */
    /*public void load(){
        File file =  new File(getNombreArchivo());
        FileInputStream fileEntrada = new FileInputStream(file);
        try {
            fileEntrada.close();
        } catch (IOException ex) {
            Logger.getLogger(Libro.class.getName()).log(Level.SEVERE, null, ex);
        }
        SXSSFWorkbook wb = new SXSSFWorkbook(fileEntrada);
        for(int i = 0; i < wb.getSheets(); i++){
            Sheet sh = wb.getSheetAt(i);
            int filas = sh.getRows();
            int columnas = 0;
            for(int j = 0; j < filas; j++){
                if(columnas < sh.getCells()){
                    columnas = sh.getCells();
                }
                Hoja hoja = new Hoja(sh.getName,filas,columnas);
                for(int k = 0; k < filas; k++){
                    Row row = sh.getRow(k);
                    for(int z = 0; z < columnas; z++){
                        Cell cell = sh.getCell(z);
                        switch (cell.getCellType()){
                            case STRING: 
                                hoja.setDato(cell.getStringCellValue(),k,z);
                                break;
                            case NUMERIC:
                                hoja.setDato(cell.getStringCellValue(),k,z);
                                break;
                            case FORMULA:
                                hoja.setDato(cell.getStringCellValue(),k,z);
                                break;
                            case BOOLEAN:
                                hoja.setDato(cell.getStringCellValue(),k,z);
                                break;
                            default:
                                hoja.setDato("",i,j);
                        }
                    }
                }
            }
        }
    }*/
    
    /**
     * Metodo load al que se le pasa el nombre del archivo
     * @param fileName 
     */
    public void load(String fileName){
        this.nombreArchivo = fileName;
        //this.load();
    }
    
    /**
     * Metodo guardar, cuya funcion es crear una hoja y guardarla
     * @throws ExcelAPIException 
     */
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
    
    /**
     * Metodo save al que se le pasa el nombre del archivo
     * @param fileName
     * @throws ExcelAPIException 
     */
    public void save(String fileName) throws ExcelAPIException{
        this.nombreArchivo = fileName; 
        this.save();
    }
    
    /**
     * Metodo que comprueba que la esxtension del fichero sea .xlsx
     */
    private void testExtension(){
        if(!this.nombreArchivo.matches("^(?:[\\w]\\:|\\\\)(\\[a-z_\\-\\s0-9\\.]+)+\\.(txt|gif|pdf|doc|docx|xls|xlsx)$")){
            this.nombreArchivo=this.nombreArchivo+".xlsx";
        }
    }
}