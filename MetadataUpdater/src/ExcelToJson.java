/**
 * Created by priyanp on 11/08/17.
 */

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;


public class ExcelToJson {

    private Workbook workbook;
    private Sheet sheet;
    private JSONObject object = new JSONObject();
    private JSONArray data = new JSONArray();
    private String inputFile;
    private String outputFile;

    public ExcelToJson(String inputFile){
        this.inputFile = inputFile;
    }

    public ExcelToJson() {

    }

    private void setInputFile() throws IOException, BiffException {
        File inputWorkbook = new File(inputFile);

        workbook = Workbook.getWorkbook(inputWorkbook);
        // Get the Metadata sheet
        sheet = workbook.getSheet("Metadata");

        String fileName;
        if(inputFile.matches("^.*[0-9]{14}.xls$")){
            fileName = inputFile.replaceFirst(".xls", "");
        }
        else{
            fileName = inputFile.replaceFirst(".xls", "") + "_" + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());;
        }
        outputFile = fileName + ".txt";
    }

    private void closeInputFile() throws IOException, WriteException {
            workbook.close();
    }


    private void readExcelFile() throws JSONException {


            for (int i = 1; i < sheet.getRows(); i++) {
                Map<String, Object> mapColNames = new HashMap<String, Object>();
                for (int j = 0; j < sheet.getColumns(); j++) {
                    Cell cell = sheet.getCell(j, i);
                    if(cell.getContents().toString()!="") {
                        if (getColumnType(sheet.getCell(j, 0).getContents()).equalsIgnoreCase("Number")) {
                            mapColNames.put(sheet.getCell(j, 0).getContents(), Integer.parseInt(cell.getContents()));
                        } else if (getColumnType(sheet.getCell(j, 0).getContents()).equalsIgnoreCase("Bool")) {
                            mapColNames.put(sheet.getCell(j, 0).getContents(), cell.getContents().equals("") ? "" : java.lang.Boolean.valueOf(cell.getContents()));
                        } else {
                            mapColNames.put(sheet.getCell(j, 0).getContents(), cell.getContents());
                        }
                    }
                }
                appendToJSON(mapColNames);
            }
    }

    private String getColumnType(String columnName){
        String columnType = "";
        Sheet sheet = workbook.getSheet("Datatype");
        for(int col = 0; col<sheet.getColumns(); col++){
            if(sheet.getCell(col,1).getContents().equalsIgnoreCase(columnName)){
                columnType = sheet.getCell(col,2).getContents();
                break;
            }
        }
        return columnType;
    }


    public void excelToJsonProcess() throws JSONException, IOException, WriteException, BiffException {
        setInputFile();
        readExcelFile();
        createJSON();
        closeInputFile();
    }

    private void appendToJSON(Map<String, Object> mapColNames) throws JSONException {
        JSONObject dataElement = new JSONObject();
        dataElement.put("columnMetadata", constructJSONElement(mapColNames, "columnMetadata"));
        dataElement.put("userPreference", constructJSONElement(mapColNames, "userPreference"));
        data.put(dataElement);
    }


    private JSONObject constructJSONElement(Map<String, Object> mapColNames , String nodeName) throws JSONException {

        JSONObject node = new JSONObject();
        JSONObject childObj;

        JSONObject parentObj = new JSONObject();

        for (Map.Entry<String, Object> keyValue : mapColNames.entrySet()) {
            Object nodeValue = keyValue.getValue();
            if (!nodeValue.toString().equalsIgnoreCase("")) {

               if (keyValue.getKey().contains(nodeName)) {
                    Object newValue;
                    if(keyValue.getValue().toString().contains(",")){
                        newValue = new JSONArray(keyValue.getValue().toString().split(","));
                    }else{
                        newValue = keyValue.getValue();
                    }
                    keyValue.setValue(newValue);
                   node = addNodes(keyValue,node);
                }
            }
        }

        return node;
    }

    private void createJSON() throws JSONException, IOException {
        object.put("Data", data);
        Writer out = new StringWriter();
        object.write(out);
        String jsonText = out.toString();
        out.close();

        File file = new File(outputFile);
        if (file.exists())
            file.delete();

        try (PrintStream out1 = new PrintStream(new FileOutputStream(outputFile))) {
            out1.print(jsonText);
        }
        JOptionPane.showMessageDialog(new Frame(),"New metadata json created: " + outputFile);
        }

    private JSONObject addNodes(Map.Entry<String, Object> keyValue, JSONObject node) throws JSONException {
            String arrKeys[] = keyValue.getKey().split("__");
            if(arrKeys.length>2) {
                String newkey = keyValue.getKey().toString();
                Map<String, Object> mapNew = new HashMap<String, Object>();
                mapNew.put(newkey.replaceFirst("__" + arrKeys[1], ""), keyValue.getValue());
                if (!node.has(arrKeys[1])) {
                    node.put(arrKeys[1], new JSONObject());
                }
                addNodes(mapNew.entrySet().iterator().next(), node.getJSONObject(arrKeys[1]));
            }
            else{
                node.put(arrKeys[1],keyValue.getValue());
            }
            return node;
        }
}
