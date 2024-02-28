/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package risingsun;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class RisingSun {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\hdzdi\\OneDrive\\Documents\\ITO\\S6\\Métodos\\examen 2DA unidad octavio MA_102115 7 MIN 100.xlsx"; // Ajusta esto a la ubicación de tu archivo Excel
        try (FileInputStream fileInputStream = new FileInputStream(filePath)) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheet("METODO GRAFICO"); // Accede a la hoja por su nombre
            
            // Modificar la celda K1
            Row row = sheet.getRow(0); // K1 está en la primera fila, índice 0
            if (row == null) {
                row = sheet.createRow(0); // Crea la fila si no existe
            }
            Cell cell = row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // La columna K es la 11ª columna, índice 10
            cell.setCellValue("2"); // Establece el nuevo valor en la celda K1
            
            // Guardar los cambios en el archivo
            try (FileOutputStream fileOutputStream = new FileOutputStream(filePath)) {
                workbook.write(fileOutputStream);
            }
            
            workbook.close(); // Cierra el workbook para liberar recursos
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
