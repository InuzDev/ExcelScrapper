package com.inudev.excelclassifier;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

import java.io.*;
import java.math.BigInteger;
import java.util.*;

public class App {
    public static void main(String[] args) {
        String excelPath = "students.xlsx"; // Excel file in project root
        String outputDir = "OUTPUT-FOLDER"; // Folder to store the Word files

        // Create output folder if it doesn't exist
        File folder = new File(outputDir);
        if (!folder.exists()) {
            folder.mkdirs();
            System.out.println("📁 Created folder: " + folder.getAbsolutePath());
        }

        try (FileInputStream fis = new FileInputStream(excelPath);
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            // Header row
            Row headerRow = rowIterator.next();
            int nameCol = -1, idCol = -1, teacherCol = -1, classCol = -1;

            for (Cell cell : headerRow) {
                String val = cell.getStringCellValue().trim().toLowerCase();
                if (val.contains("nombre"))
                    nameCol = cell.getColumnIndex();
                else if (val.equals("id"))
                    idCol = cell.getColumnIndex();
                else if (val.contains("profesor"))
                    teacherCol = cell.getColumnIndex();
                else if (val.contains("clase"))
                    classCol = cell.getColumnIndex();
            }

            // Group students by teacher + class
            Map<String, List<String[]>> teacherClassStudents = new HashMap<>();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (row.getCell(teacherCol) == null || row.getCell(classCol) == null)
                    continue;

                String name = row.getCell(nameCol).getStringCellValue().trim();

                // Read ID safely
                Cell idCell = row.getCell(idCol);
                String id = "";
                if (idCell != null) {
                    switch (idCell.getCellType()) {
                        case STRING:
                            id = idCell.getStringCellValue().trim();
                            break;
                        case NUMERIC:
                            double numericValue = idCell.getNumericCellValue();
                            if (numericValue == Math.floor(numericValue))
                                id = String.valueOf((long) numericValue);
                            else
                                id = String.valueOf(numericValue);
                            break;
                        default:
                            id = idCell.toString().trim();
                            break;
                    }
                }

                String teacher = row.getCell(teacherCol).getStringCellValue().trim();
                String className = row.getCell(classCol).getStringCellValue().trim();

                String key = teacher + " | " + className;

                teacherClassStudents
                        .computeIfAbsent(key, k -> new ArrayList<>())
                        .add(new String[] { name, id });
            }

            // === Generate documents ===
            for (String key : teacherClassStudents.keySet()) {
                String[] parts = key.split("\\|");
                String teacher = parts[0].trim();
                String className = parts[1].trim();

                XWPFDocument doc = new XWPFDocument();

                // === Safe page margins ===
                int margin = 720; // ~1 inch
                if (doc.getDocument().getBody().getSectPr() == null) {
                    doc.getDocument().getBody().addNewSectPr();
                }
                CTPageMar pgMar = doc.getDocument().getBody().getSectPr().addNewPgMar();
                pgMar.setLeft(BigInteger.valueOf(margin));
                pgMar.setRight(BigInteger.valueOf(margin));
                pgMar.setTop(BigInteger.valueOf(margin));
                pgMar.setBottom(BigInteger.valueOf(margin));

                // === Title ===
                XWPFParagraph title = doc.createParagraph();
                title.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun run = title.createRun();
                run.setBold(true);
                run.setFontSize(16);
                run.setText("Profesor: " + teacher);
                run.addBreak();
                if (className == "" || className == " ") {
                    className = "Sin mencionar";
                }
                run.setText("Clase: " + className);
                run.addBreak();

                // === Table ===
                XWPFTable table = doc.createTable();

                // Set clean borders that Word always accepts
                table.getCTTbl().getTblPr().unsetTblBorders();
                CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
                borders.addNewInsideH().setVal(STBorder.SINGLE);
                borders.addNewInsideV().setVal(STBorder.SINGLE);
                borders.addNewLeft().setVal(STBorder.SINGLE);
                borders.addNewRight().setVal(STBorder.SINGLE);
                borders.addNewTop().setVal(STBorder.SINGLE);
                borders.addNewBottom().setVal(STBorder.SINGLE);

                // Column widths (fixed)
                int[] colWidths = { 4000, 2000, 3000 };
                XWPFTableRow header = table.getRow(0);

                // Create headers manually
                String[] headers = { "Nombre", "ID", "Clase" };
                for (int i = 0; i < headers.length; i++) {
                    XWPFTableCell cell = (i == 0) ? header.getCell(0) : header.addNewTableCell();
                    cell.removeParagraph(0);
                    XWPFParagraph p = cell.addParagraph();
                    p.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun r = p.createRun();
                    r.setBold(true);
                    r.setFontSize(12);
                    r.setText(headers[i]);
                    cell.setColor("D9D9D9");
                    cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(colWidths[i]));
                }

                // === Fill table ===
                for (String[] info : teacherClassStudents.get(key)) {
                    XWPFTableRow row = table.createRow();
                    row.getCell(0).setText(info[0]);
                    row.getCell(1).setText(info[1]);
                    row.getCell(2).setText(className);
                }

                for (XWPFTableRow row : table.getRows()) {
                    row.setHeight(400);
                }

                // Save file
                String safeTeacher = teacher;
                String safeClass = className;
                if (safeClass == "" || safeClass == " ") {
                    safeClass = "Sin mencionar";
                }
                String fileName = safeTeacher + " - " + safeClass + " - estudiantes.docx";
                File outFile = new File(folder, fileName);

                try (FileOutputStream out = new FileOutputStream(outFile)) {
                    doc.write(out);
                }

                doc.close();
                System.out.println("Created: " + outFile.getPath());
            }

            System.out.println("\nAll Word documents have been created successfully in: " + folder.getAbsolutePath());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
