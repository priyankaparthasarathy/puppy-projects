/**
 * Created by priyanp on 11/08/17.
 */

import jxl.Cell;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Boolean;
import jxl.write.Label;
import jxl.write.Number;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.*;


public class JsonToExcel {

    private String inputFile;
    private String outputFile;
    private WritableWorkbook workbook;
    private WritableSheet sheet;
    private WritableSheet datatype_sheet;


    public JsonToExcel(String inputFile){
        this.inputFile = inputFile;
    }

    public JsonToExcel() {

    }

    private void createExcelFile() throws IOException {
        String fileName;
        if(inputFile.matches("^.*[0-9]{14}.txt$")){
            fileName = inputFile.replaceFirst(".txt", "").replaceFirst(".json", "");

        }
        else{
             fileName = inputFile.replaceFirst(".txt", "").replaceFirst("json", "") + "_" + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());;
        }

       outputFile = fileName + ".xls";

        File file = new File(outputFile);
        if (file.exists())
            file.delete();

        workbook = Workbook.createWorkbook(file);
        sheet = workbook.createSheet("Metadata", 0);
        datatype_sheet = workbook.createSheet("Datatype", 1);
        datatype_sheet.getSettings().setHidden(true);
    }

    private void addColumns(Map<String, Object> mapColNames) throws WriteException, IOException {

        int intRows = sheet.getRows();
        intRows = intRows == 0 ? intRows = 1 : intRows;

        for(Map.Entry<String, Object> addCol:mapColNames.entrySet()) {
            int intColNum = -1;
            int intCols = sheet.getColumns();

            Label label;
            Number numberCell;
            Boolean boolCell;

            for (int col = 0; col < intCols; col++) {
                Cell header = sheet.getCell(col, 0);

                if (header.getContents().equalsIgnoreCase(addCol.getKey().toString())) {
                    intColNum = col;
                    break;
                }
            }

            if (intColNum < 0) {
                label = new Label(intCols, 0, addCol.getKey().toString());
                sheet.addCell(label);
                datatype_sheet.addCell(new Label(intCols, 1, addCol.getKey().toString()));
                datatype_sheet.addCell(new Label(intCols, 2, ""));
                intColNum = intCols;
            }

            if (addCol.getValue() instanceof Integer){

                numberCell = new Number(intColNum, intRows, ((Integer) addCol.getValue()));
                sheet.addCell(numberCell);
                ((Label)datatype_sheet.getWritableCell(intColNum, 2)).setString("Number");
            }else{
                label = new Label(intColNum, intRows, addCol.getValue().toString());
                sheet.addCell(label);
                if(addCol.getValue().toString().equalsIgnoreCase("true") || addCol.getValue().toString().equalsIgnoreCase("false")){
                    ((Label) datatype_sheet.getWritableCell(intColNum, 2)).setString("Bool");
                }else {
                    ((Label) datatype_sheet.getWritableCell(intColNum, 2)).setString("String");
                }
            }
        }

    }

    private void closeExcelFile()
            throws WriteException, IOException {
        workbook.write();
        workbook.close();
        JOptionPane.showMessageDialog(new Frame(),"Metadata converted to excel: " + outputFile);

    }


    private String readJSONFile() throws IOException {
        InputStream is = new FileInputStream(inputFile);
        BufferedReader buf = new BufferedReader(new InputStreamReader(is));

        String line = buf.readLine();
        StringBuilder sb = new StringBuilder();

        while(line != null){
            sb.append(line).append("\n");
            line = buf.readLine();
        }

        String fileAsString = sb.toString();
        return fileAsString;

    }

    public void jsonToExcelProcess() throws JSONException, IOException, WriteException {
        createExcelFile();
        parseJSON();
        closeExcelFile();
    }


    private void parseJSON() throws IOException, JSONException, WriteException {
        String JSON_DATA = readJSONFile();

        final JSONObject obj = new JSONObject(JSON_DATA);
        final JSONArray metadata = obj.getJSONArray("Data");
        final int n = metadata.length();


        for (int i = 0; i < n; i++) {
            Map<String, Object> mapColNames = new HashMap<String, Object>();
            JSONObject colData = metadata.getJSONObject(i);

            if (colData.has("columnMetadata")) {
                JSONObject columnMetadata = colData.getJSONObject("columnMetadata");
                extractObjects(mapColNames, columnMetadata, "columnMetadata");
            }
            if(colData.has("userPreference")) {
                JSONObject userPreference = colData.getJSONObject("userPreference");
                extractObjects(mapColNames, userPreference, "userPreference");
            }
            addColumns(mapColNames);
        }


    }


    private void extractObjects(Map<String, Object> mapColNames, JSONObject jsonObj, String prefix) throws JSONException, IOException, WriteException {
        if(!prefix.endsWith("__") && prefix!="")
            prefix+="__";

        for (int j = 0; j < jsonObj.names().length(); j++) {
            Object key = jsonObj.names().get(j);

            String strKey = key.toString();
            Object strVal = jsonObj.get(strKey);
            Object strValNew = new Object();

            if (strVal instanceof JSONObject) {
                if(prefix=="")
                    extractObjects(mapColNames, (JSONObject) strVal, strKey + "__");
                else
                    extractObjects(mapColNames, (JSONObject) strVal, prefix + strKey + "__");
            } else {
                if(strVal instanceof JSONArray){
                    strValNew = ((JSONArray) strVal).get(0);
                    for(int arrCounter = 1; arrCounter<((JSONArray) strVal).length(); arrCounter++)
                        strValNew = strValNew + "," + ((JSONArray) strVal).get(arrCounter) ;
                    strVal = strValNew;
                }
                mapColNames.put(prefix + strKey, strVal);

            }

        }
    }

}
