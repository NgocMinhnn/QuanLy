/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package BUS;

import java.sql.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;

/**
 *
 * @author phuon
 */
public class EditString {
    public static String DateToString(Date date){
        String dateString="";
        DateFormat df = new SimpleDateFormat("yyyy/MM/dd");
        dateString = df.format(date);
        return dateString;
    }
    public static void main(String[] args) {
        LocalDate locald = LocalDate.of(2001, 10, 20);
        Date date = Date.valueOf(locald); // Magic happens here!
        //r.setDateOfBirth(date);
        System.out.println(DateToString(date));
    }
}
