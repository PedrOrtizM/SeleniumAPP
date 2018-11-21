package com.accenture.retoExcel;

import java.io.File;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.accenture.retoExcel.User;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
  
    	//--------------------------------------| leer archivo excel | ---------------------------------------
    	
    	ArrayList<String> emails = new ArrayList<String>();
        try {
        	
        	// Abrimos en libro con la ubicación del archivo
			Workbook libro = WorkbookFactory.create(new File("C:/Users/administrador/Documents/emails.xlsx"));
			//obtenemos la primera hoja del libro
			Sheet hoja = libro.getSheetAt(0);
			DataFormatter dataFormatter = new DataFormatter();
		
		    //recorremos las filas de la hoja
			for (Row row: hoja) {
		        	//recorremos las celdas de la hoja
		           for(Cell celda: row) {
		        	   //se obtiene el valor de la celda 
		                String cellValue = dataFormatter.formatCellValue(celda);
		                // agregamos al arraylist el valor 
		                emails.add(cellValue);
		       //        System.out.print(cellValue);
		            }
		       //     System.out.println();
		        }
			//cerramos el libro
		        libro.close();
			
	
		} catch (Exception e) {
			
			System.out.println("error "+e.toString());
		}
        
  //--------------------------------------| automatización | ---------------------------------------
        
        
        User usuario = new User();
        WebDriver web = new ChromeDriver();
        //Abre google
        web.get("http://www.google.com");
        // click en Gmail
        web.findElement(By.xpath("//*[@id=\"gbw\"]/div/div/div[1]/div[1]/a")).click();
        //click en login
        web.findElement(By.xpath("/html/body/nav/div/a[2]")).click();
        // Introduce email
        web.findElement(By.xpath("//*[@id=\"identifierId\"]")).sendKeys(usuario.getEmail()+"\n");
        esperar(2);
        // Introduce contraseña 
        web.findElement(By.xpath("//*[@id=\"password\"]/div[1]/div/div[1]/input")).sendKeys(usuario.getPass());
        esperar(2);
        // Click en entrar 
        web.findElement(By.xpath("//*[@id=\"passwordNext\"]/content/span")).click();
        esperar(5);
        // Click en redactar 
        web.findElement(By.xpath("//*[@id=\":ir\"]/div/div")).click();        
        esperar(3);
        // Recorre los emails y agrega a destinatario
        emails.forEach((t)->{
        	 web.findElement(By.xpath("//textarea[@name='to']")).sendKeys(t+"\n");
             esperar(3);
         });
        // Agrega Asunto
         web.findElement(By.xpath("//input[@name='subjectbox']")).sendKeys("Prueba Selenium");
         // Agrega un mensaje
         web.findElement(By.xpath("//*[@id=\":ou\"]")).sendKeys("Mensaje de prueba automática");
         // Click en enviar
         web.findElement(By.xpath("//*[@id=\":nf\"]")).click();
         esperar(2);
         // web.quit();
         
         System.out.println("Mensaje enviado con éxito");

       
    }
    
    public static void esperar(int segundos){
        try {
    			Thread.sleep(segundos*1000);
    		} catch (InterruptedException e) {
    			
    			System.out.println("Error en el hilo"+e.toString());
    		}
            
    	
    }
}
