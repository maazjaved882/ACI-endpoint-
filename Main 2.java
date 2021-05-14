import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.*;
import java.util.*;

public class Main {

    public static void main(String[] args) {
        String fileName = "D:/test.json";
        JSONParser jsonParser = new org.json.simple.parser.JSONParser();

        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet");
        String[] headers = {"Tenant","Application" ,"EPG",    "Mac address", "IP address"   ,"Node 1","Node 2","Interfaces" };
        try (FileReader reader = new FileReader(fileName)) {
            Row row = sheet.createRow(0);
            int i=0;
            for (String header : headers) {
                Cell cell = row.createCell(i++);
                cell.setCellValue(header);
            }

            JSONObject obj = (JSONObject) jsonParser.parse(reader);
            JSONArray rows = ((JSONArray)obj.get("imdata"));
            for( i=0;i<rows.size(); i++) {
                JSONObject attributes = (JSONObject) ((JSONObject)((JSONObject)rows.get(i)).get("fvCEp")).get("attributes");
                String dn = attributes.get("dn").toString();
                String[] splits = dn.split("/");

                row = sheet.createRow(i+1);
                if(splits[1].startsWith("tn-")) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue(splits[1].substring(3));
                }

                if(splits[2].startsWith("ap-")) {
                    Cell cell = row.createCell(1);
                    cell.setCellValue(splits[2].substring(3));
                }

                if(splits[3].startsWith("epg-")) {
                    Cell cell = row.createCell(2);
                    cell.setCellValue(splits[3].substring(4));
                }

                Cell cell = row.createCell(3);
                cell.setCellValue(attributes.getOrDefault("mac","").toString());

                cell = row.createCell(4);
                cell.setCellValue(attributes.getOrDefault("ip","").toString());

                cell = row.createCell(7);
                cell.setCellValue(attributes.getOrDefault("encap","").toString());

                JSONArray children = (JSONArray) ((JSONObject)((JSONObject)rows.get(i)).get("fvCEp")).get("children");

                JSONObject interfacesObject = null;

                if(children!=null) {
                    for(int j = 0; j< children.size(); j++) {

                        if(((JSONObject)children.get(j)).containsKey("fvRsCEpToPathEp")) {
                            interfacesObject = (JSONObject)((JSONObject)((JSONObject)children.get(j)).get("fvRsCEpToPathEp")).get("attributes");
                            if(interfacesObject.get("rn") != null && interfacesObject.get("rn").toString().contains("pathep-")) {
                                break;
                            }
                        }
                    }
                }



                if(interfacesObject!=null) {
                    String rn = interfacesObject.get("rn").toString();

                    if(rn.contains("protpaths-")) {
                        int protoIndex = rn.indexOf("protpaths-");
                        int endIndex = protoIndex;

                        for(int j = protoIndex ; j< rn.length() ; j++) {
                            if(rn.charAt(j) == '/') {
                                endIndex = j;
                                break;
                            }
                        }

                        if(protoIndex!= endIndex) {
                            String protPaths = rn.substring(protoIndex, endIndex);
                            cell = row.createCell(5);
                            cell.setCellValue(protPaths.split("-")[1]);

                            cell = row.createCell(6);
                            cell.setCellValue(protPaths.split("-")[2]);

                        }
                    } else if(rn.contains("/paths-")) {
                        int protoIndex = rn.indexOf("paths-");
                        int endIndex = protoIndex;

                        for(int j = protoIndex ; j< rn.length() ; j++) {
                            if(rn.charAt(j) == '/') {
                                endIndex = j;
                                break;
                            }
                        }

                        if(protoIndex!= endIndex) {
                            String protPaths = rn.substring(protoIndex, endIndex);
                            cell = row.createCell(5);
                            cell.setCellValue(protPaths.split("-")[1]);

                        }
                    }

                    if(rn.contains("/pathep-")) {
                        int pathEpIndex = rn.indexOf("/pathep-");
                        cell = row.createCell(7);
                        cell.setCellValue( rn.substring(pathEpIndex + 9, rn.length()-2) );
                    }
                }

                //System.out.println(splits[1].substring(3));
            }
            //JSONObject attributes = (JSONObject) ((JSONObject)((JSONObject)((JSONArray)obj.get("imdata")).get(0)).get("fvCEp")).get("attributes");
            for ( i=0 ; i<8 ; i++) {
                sheet.autoSizeColumn(i);
            }

            try (OutputStream fileOut = new FileOutputStream("workbook.xls")) {
                wb.write(fileOut);
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        }
    }
}