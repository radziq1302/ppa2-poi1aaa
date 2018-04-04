package sample;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import javafx.scene.paint.Color;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import javafx.event.ActionEvent;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;

/**
 * Created by pwilkin on 30-Nov-17.
 */
public class DaneOsobowe implements HierarchicalController<MainController> {

    public TextField imie;
    public TextField nazwisko;
    public TextField pesel;
    public TextField indeks;
    public TableView<sample.Student> tabelka;
    private sample.MainController parentController;

    public void dodaj(ActionEvent actionEvent) {
        sample.Student st = new sample.Student();
        st.setName(imie.getText());
        st.setSurname(nazwisko.getText());
        st.setPesel(pesel.getText());
        st.setIdx(indeks.getText());
        tabelka.getItems().add(st);
    }

    public void setParentController(sample.MainController parentController) {
        this.parentController = parentController;
        //tabelka.getItems().addAll(parentController.getDataContainer().getStudents());
        tabelka.setItems(parentController.getDataContainer().getStudents());
    }

    public void usunZmiany() {
        tabelka.getItems().clear();
        tabelka.getItems().addAll(parentController.getDataContainer().getStudents());
    }

    public sample.MainController getParentController() {
        return parentController;
    }

    public void initialize() {
        for (TableColumn<sample.Student, ?> studentTableColumn : tabelka.getColumns()) {
            if ("imie".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("name"));
            } else if ("nazwisko".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("surname"));
            } else if ("pesel".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("pesel"));
            } else if ("indeks".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("idx"));
            }
        }

    }

    public void zapisz(ActionEvent actionEvent) {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("Studenci");

        XSSFCellStyle style = wb.createCellStyle();
        XSSFCellStyle style1 = wb.createCellStyle();
        XSSFCellStyle style2 = wb.createCellStyle();
        XSSFCellStyle style3 = wb.createCellStyle();
        int row = 1;
        XSSFFont font= wb.createFont();

        font.setColor(HSSFFont.COLOR_RED);
        font.setBold(true);
        XSSFRow r1 = sheet.createRow(row);
 //       style=r1.getRowStyle();
        style1.setFillForegroundColor(HSSFFont.COLOR_RED);
        XSSFColor myColor = new XSSFColor(java.awt.Color.YELLOW);
        style2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        XSSFColor myColor1 = new XSSFColor(java.awt.Color.GREEN);
        style3.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(font);
        String [] lista = {"Imię", "Nazwisko", "Ocena", "Szczegóły", "Indeks", "Pesel"};
        XSSFRow init = sheet.createRow(0);
        int ce = 0;
        for (String s : lista) {
            XSSFCell ci = init.createCell(ce);//
            ci.setCellStyle(style);
            ci.setCellValue(s);
            ce++;
        }
        /*r1.setRowStyle(style);

        XSSFCell c = r1.createCell(0);//
        c.setCellStyle(style);
        c.setCellValue("Imię");
        r1.createCell(1).setCellValue("Nazwisko");
        r1.createCell(2).setCellValue("Ocena");

        r1.createCell(3).setCellValue("szczegóły");
        r1.createCell(4).setCellValue("Indeks");
        r1.createCell(5).setCellValue("Pesel");
        style=r1.getRowStyle();
        style.setFont(font);
        r1.setRowStyle(style);
        row++;*/

        // sprawdz czy dziala
        for (sample.Student student : tabelka.getItems()) {
            XSSFRow r = sheet.createRow(row);
            r.createCell(0).setCellValue(student.getName());
            r.createCell(1).setCellValue(student.getSurname());
            XSSFCell ciel = null;

            Double ocen = student.getGrade();
            if (ocen != null) {

                if (ocen < 3.0) {

                    ciel = r.createCell(2);
                    ciel.setCellStyle(style2);
                    ciel.setCellValue(student.getGrade());
                }
                if (ocen >= 3.0) {
                    System.out.println("to jest cos" + student.getGrade());
                    ciel = r.createCell(2);
                    ciel.setCellStyle(style3);
                    ciel.setCellValue(student.getGrade());
                }

                //ciel.setCellValue(0);
            }
            else if (ocen == null) {
                ciel = r.createCell(2);
                ciel.setCellStyle(style1); }
            // tutaj kolorki
            r.createCell(3).setCellValue(student.getGradeDetailed());
            r.createCell(4).setCellValue(student.getIdx());
            r.createCell(5).setCellValue(student.getPesel());
            row++;
        }

        try (FileOutputStream fos = new FileOutputStream("data.xlsx")) {
            wb.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /** Uwaga na serializację: https://sekurak.pl/java-vs-deserializacja-niezaufanych-danych-i-zdalne-wykonanie-kodu-czesc-i/ */
    public void wczytaj(ActionEvent actionEvent) {
        ArrayList<sample.Student> studentsList = new ArrayList<>();
        try (FileInputStream ois = new FileInputStream("data.xlsx")) {
            XSSFWorkbook wb = new XSSFWorkbook(ois);
            XSSFSheet sheet = wb.getSheet("Studenci");
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                XSSFRow row = sheet.getRow(i);

                sample.Student student = new sample.Student();
                student.setName(row.getCell(0).getStringCellValue());
                student.setSurname(row.getCell(1).getStringCellValue());
                if (row.getCell(2) != null) {
                    student.setGrade(row.getCell(2).getNumericCellValue());
                }
                if (row.getCell(3) != null) {
                student.setGradeDetailed(row.getCell(3).getStringCellValue()); }
                student.setIdx(row.getCell(4).getStringCellValue());
                student.setPesel(row.getCell(5).getStringCellValue());
                studentsList.add(student);
            }
            tabelka.getItems().clear();
            tabelka.getItems().addAll(studentsList);
            ois.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
