/**
 * Created by priyanp on 11/08/17.
 */

import javax.swing.*;
import java.awt.*;
import java.io.*;


public class Main extends JFrame {



    public static void main(String args[]) throws IOException {

        Main m = new Main();
        String ret;
        int choice = JOptionPane.showOptionDialog(null,
                "What would you like to convert?",
                "Metadata Updater",
                JOptionPane.OK_CANCEL_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                new String[]{"Cancel", "Excel to JSON" , "JSON to Excel"},
                "JSON to Excel");
        if(choice == 2 ){
            ret = m.showFileDialog(new Frame(), "Choose File", "", "txt");
            if(ret!=null) {
                if(!ret.endsWith(".txt") && !ret.endsWith(".json")){
                    JOptionPane.showMessageDialog(null, "Invalid file type. Please choose a '.txt' or '.json' file", "Error", JOptionPane.ERROR_MESSAGE);
                }else {
                    JsonToExcel j2e = new JsonToExcel(ret);
                    try {
                        j2e.jsonToExcelProcess();
                    } catch (Exception e) {
                        JOptionPane.showMessageDialog(null, "Error converting JSON to Excel", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
        }else if(choice==1){
            ret = m.showFileDialog(new Frame(), "Choose File", "", "xls");
            if(ret!=null) {
                if(!ret.endsWith(".xls")){
                    JOptionPane.showMessageDialog(null, "Invalid file type. Please choose a '.xls' file", "Error", JOptionPane.ERROR_MESSAGE);
                }else {

                    ExcelToJson e2j = new ExcelToJson(ret);
                    try {
                        e2j.excelToJsonProcess();
                    } catch (Exception e) {
                        JOptionPane.showMessageDialog(null, "Error converting Excel to Json", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
        }

        System.exit(0);

    }


    String showFileDialog (Frame frame, String dialogTitle, String defaultDirectory, String fileType)
    {
        FileDialog fd = new FileDialog(frame, dialogTitle, FileDialog.LOAD);
        fd.setFile(fileType);
        fd.setDirectory(defaultDirectory);
        fd.setLocationRelativeTo(frame);
        fd.setVisible(true);

        String directory = fd.getDirectory();
        String filename = fd.getFile();
        if (directory == null || filename == null || directory.trim().equals("") || filename.trim().equals(""))
        {
            return null;
        }
        else
        {
            return directory + filename;
        }
    }

}
