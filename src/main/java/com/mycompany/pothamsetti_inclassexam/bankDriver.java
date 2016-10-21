/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.pothamsetti_inclassexam;

import com.itextpdf.text.BadElementException;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 *
 * @author krishna jyothi swaroop reddy pothamsetti
 */
public class bankDriver {

    public static void main(String args[]) throws FileNotFoundException, DocumentException, BadElementException, IOException {
        ReadingInputFromExcel rd = new ReadingInputFromExcel();
        List<Bank> customerList = rd.getCustomerListFromExcel();
        for (Bank x : customerList) {
            try (OutputStream file = new FileOutputStream(new File(x.getLastName() + ".pdf"))) {
                Document document = new Document();
                PdfWriter.getInstance(document, file);

                //Inserting Image in PDF
                Image image = Image.getInstance("logo.jpg");
                //Setting Image Width and Height
                image.scaleAbsolute(500, 50f);

                //Now Insert Every Thing Into PDF Document
                document.open();//PDF document opened........			       
                document.add(image);
                document.add(new Paragraph("----------------------------------------------------------------------------------------------------------------------------------"));
                document.add(new Paragraph("   "));
                document.add(new Paragraph("   "));
                document.add(new Paragraph("Welcome! " + x.getFirstName() + " " + x.getLastName() + "!"));
                document.add(new Paragraph(" "));
                document.add(new Paragraph("        " + "Below are your Account Details :"));
                document.add(new Paragraph("        " + "First Name: " + x.getFirstName()));
                document.add(new Paragraph("        " + "Last Name: " + x.getLastName()));
                document.add(new Paragraph("        " + "Account Number: " + x.getAccountNumber()));
                document.add(new Paragraph("        " + "Account Balance: $" + x.getAccountBalance()));
                document.close();
            }
        }
    }

}
