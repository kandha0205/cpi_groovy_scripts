package com.test.groovy.Tester

import com.sap.gateway.ip.core.customdev.util.Message
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.text.ParseException
import java.text.SimpleDateFormat;



def input = "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"

def out = convert(input)
println()
println(out)


def convert(String input) {
// logging



    byte[] data = java.util.Base64.getDecoder().decode(input);

    def is = new ByteArrayInputStream(data);

   // def input = new FileInputStream(new File("C:\\Users\\M000077\\OneDrive - Uniper SE\\CPI\\uniper\\English_Receive_Single_Line.xlsx"))
    def output = convertExcelToCSV(is)

    return  output

}


def String convertExcelToCSV(InputStream is) throws Exception {
    StringBuilder sb = new StringBuilder();


    try {

        int sendIndex;
        int receiveIndex;
        int messageTimeIndex;
        int nameIndex;
        int referencetimeBeginIndex;
        int referencetimeEndIndex;
        int revNoIndex;
        int statusIndex;
        int receiveSearchTerm;
        int receivesenderIndex;
// private static boolean isGerman = false;
        String isSend
        String LINE_FEED
        String send_or_receive_condion
        String fileInput
        boolean isRunningInEclipse
        int count;

        send_or_receive_condion = "No_MessageType_Found";

     isSend = false;
       LINE_FEED = "\r\n";


      fileInput = null;
         isRunningInEclipse = false;
        count;
        count = 0


        Workbook workbook = WorkbookFactory.create(is);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(6);
        // System.out.println(row.getLastCellNum());

        for (int i = 0; i < row.getLastCellNum(); i++) {
            String cellValue = row.getCell(i).toString();
            if (cellValue != "") {
                // System.out.print("cell no."+i+" :"+row.getCell(i));

                //	System.out.println("cell no." + i + " :" + cellValue);

                switch (cellValue) {

                // Sender Specific

                    case "Send": // 1. Send - English
                        sendIndex = i;
                        isSend = true;
                        send_or_receive_condion = "Send";
                        break;


                    case "Versandzeitpunkt": // 1. send - German
                        receiveIndex = i;
                        isSend = true;
                        send_or_receive_condion = "Send";
                        break;


                    case "Message time": // 2. Send - English
                        messageTimeIndex = i;
                        break;

                    case "Nachrichtenzeitpunkt": // 2. Receive - German
                        messageTimeIndex = i;
                        break;

                    case "Name": // 3. Send - English and German
                        nameIndex = i;
                        break;

                    case "Reference time Begin": // 4. Send English
                        referencetimeBeginIndex = i;
                        break;

                    case "Referenzzeit Beginn": // 4. Send German
                        referencetimeBeginIndex = i;
                        break;

                    case "Reference time End": // 5. Send English
                        referencetimeEndIndex = i;
                        break;

                    case "Referenzzeit Ende": // 5. Send German
                        referencetimeEndIndex = i;
                        break;

                    case "Rev. no.": // 6. Send English
                        revNoIndex = i;
                        break;

                    case "Rev-Nr": // 6. Send German
                        revNoIndex = i;
                        break;

                    case "Status": // 7. Send English and German
                        statusIndex = i;
                        break;

                        // Receiver Specific

                    case "Received": // 1. Receive - English
                        receiveIndex = i;
                        send_or_receive_condion = "Receive";
                        break;

                    case "Empfangszeitpunkt": // 1. receive - German
                        receiveIndex = i;
                        send_or_receive_condion = "Receive";
                        break;

                    case "Sender": // German
                        receivesenderIndex = i;
                        break;

                    case "Sender:": // German
                        receivesenderIndex = i;
                        break;

                    case "Suchbegriff": // 1. receiver - German
                        receiveSearchTerm = i;
                        break;

                    case "Search Term": // German
                        receiveSearchTerm = i;
                        break;

                    default:
                        break;
                }
            }
        }

        print statusIndex


        if(send_or_receive_condion.equals("No_MessageType_Found"))
        {
            //System.out.println("in the error block");
          //  trace.addWarning("Error: No Message type found, check the input file colunm names");
            throw new Exception("Input Data format is not valid. Check the colunm names. valid colunm names Send, Versandzeitpunkt or Received, Empfangszeitpunkt ");
        }







        if (isSend == true) {
            // convert it to send csv format
            //System.out.println("in Sending type");

            sb.append(
                    "Versandzeitpunkt;Nachrichtenzeitpunkt;Name;Referenzzeit Beginn;Referenzzeit Ende;Rev-Nr;Status");
            for (int i = 7; i <= sheet.getLastRowNum(); i++) {
                Row new_row = sheet.getRow(i);

                print new_row.getCell(10)

                if (new_row.getCell(statusIndex).toString().equals("Send OK with ack") && (!new_row.getCell(referencetimeBeginIndex).toString().equals("") || !new_row.getCell(referencetimeEndIndex).toString().equals("")))

                {

                    count = count + 1;

                    sb.append(LINE_FEED);

                    String versandzeitpunkt = convertDate(new_row.getCell(sendIndex).toString());
                    String nachrichtenzeitpunkt = convertDate(new_row.getCell(messageTimeIndex).toString());
                    String referenzzeitBeginn = convertDate(new_row.getCell(referencetimeBeginIndex).toString());
                    String referenzzeitEnde = convertDate(new_row.getCell(referencetimeEndIndex).toString());
                    String output = versandzeitpunkt + ";" + nachrichtenzeitpunkt + ";"
                    +new_row.getCell(nameIndex).toString() + ";" + referenzzeitBeginn + ";" + referenzzeitEnde
                    +";" + new_row.getCell(revNoIndex).toString() + ";"
                    +new_row.getCell(statusIndex).toString();
                    //	 System.out.println(output);
                    // System.exit(1);
                    sb.append(output);
                    // return;
                }

            }
        } else {
            // convert it to receive csv format
            System.out.println("in receiving type");
            sb.append(
                    "Empfangszeitpunkt;Nachrichtenzeitpunkt;Name;Referenzzeit Beginn;Referenzzeit Ende;Sender;Suchbegriff;Status");
            for (int i = 7; i <= sheet.getLastRowNum(); i++) {

                Row new_row = sheet.getRow(i);

                print new_row.getCell(10)

                // Only Lines with Status Equals "Values written" should be processed
                if (new_row.getCell(statusIndex).toString().equals("Values written")) {

                    count = count + 1;

                    println sendIndex
                    println   messageTimeIndex
                    println    referencetimeBeginIndex
                    println    referencetimeEndIndex
                    println    nameIndex
                    println    receivesenderIndex
                    println     receiveSearchTerm
                    println    statusIndex
                    println "name "+ nameIndex

                    sb.append(LINE_FEED);

                    println(new_row.getCell(statusIndex))

                    def versandzeitpunkt = convertDate(new_row.getCell(sendIndex).toString());
                    def nachrichtenzeitpunkt = convertDate(new_row.getCell(messageTimeIndex).toString());
                    def referenzzeitBeginn = convertDate(new_row.getCell(referencetimeBeginIndex).toString());
                    def referenzzeitEnde = convertDate(new_row.getCell(referencetimeEndIndex).toString());
                    def name = new_row.getCell(nameIndex).toString()
                    def receiveSearchTermValue =   new_row.getCell(receiveSearchTerm).toString()
def statusIndexValue = new_row.getCell(statusIndex).toString()

                    def output = versandzeitpunkt+";"+nachrichtenzeitpunkt+";"+referenzzeitBeginn+";"+referenzzeitEnde+";"+name+";"+receiveSearchTermValue

                    //System.out.println(output);
                    // System.exit(1);
                    sb.append(output);
                    // return;
                    // System.out.println();
                }
            }
        }

        //System.out.println(sb.toString());

        //System.exit(1);
        // System.out.println(send_or_receive_condion);

        //	 System.out.println(sb.toString());

        // System.out.println(count);

        //trace.addWarning("total records"+count);

        print sb.toString()
    }
    catch (Exception e) {
        e.printStackTrace()
    }


    return  sb.toString()
}


def String convertDate(String inputDate) throws ParseException {

    // System.out.println(inputDate);

    if (inputDate.contains(".")) {
        return inputDate;
    }
    SimpleDateFormat format1 = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
    SimpleDateFormat format2 = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
    Date date = format1.parse(inputDate);
    String outputDate = format2.format(date);
    // System.out.println(outputDate.toString());
    return outputDate.toString();
}
