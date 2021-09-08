package script

import com.sap.gateway.ip.core.customdev.util.Message;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


def Message processData(Message message) {
// logging

    def messageLog = messageLogFactory.getMessageLog(message);
    if(messageLog != null)
    {
        messageLog.setStringProperty("log1","starting")
        ;
    }



    def body = message.getBody(java.lang.String)
    byte[] data = java.util.Base64.getDecoder().decode(body);
    def input = new ByteArrayInputStream(data)

    def output = convertExcelToCSV(input,messageLog)

    message.setBody(output)

    return  message

}


def String convertExcelToCSV(InputStream is, def messageLog) throws Exception {

    if(messageLog != null)
    {
        messageLog.setStringProperty("start","start")
        ;
    }
    StringBuilder sb = new StringBuilder();
    def sendIndex;
    def receiveIndex;
    def messageTimeIndex;
    def nameIndex;
    def referencetimeBeginIndex;
    def referencetimeEndIndex;
    def revNoIndex;
    def statusIndex;
    def receiveSearchTerm;
    def receivesenderIndex;
    // private static boolean isGerman = false;
    def isSend = false;
    def LINE_FEED = "\r\n";
    //private String send_or_receive_condion = "Receive";
    def send_or_receive_condion = "No_MessageType_Found";
    def fileInput = null;
    def isRunningInEclipse = false;
    def count;
    try {

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



        if(send_or_receive_condion.equals("No_MessageType_Found"))
        {
            //System.out.println("in the error block");
            trace.addWarning("Error: No Message type found, check the input file colunm names");
            throw new Exception("Input Data format is not valid. Check the colunm names. valid colunm names Send, Versandzeitpunkt or Received, Empfangszeitpunkt ");
        }







        if (isSend == true) {
            // convert it to send csv format
            //System.out.println("in Sending type");

            sb.append(
                    "Versandzeitpunkt;Nachrichtenzeitpunkt;Name;Referenzzeit Beginn;Referenzzeit Ende;Rev-Nr;Status");
            for (int i = 7; i <= sheet.getLastRowNum(); i++) {
                Row new_row = sheet.getRow(i);

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

                // Only Lines with Status Equals "Values written" should be processed
                if (new_row.getCell(statusIndex).toString().equals("Values written")) {

                    count = count + 1;

                    sb.append(LINE_FEED);
                    String versandzeitpunkt = convertDate(new_row.getCell(sendIndex).toString());
                    String nachrichtenzeitpunkt = convertDate(new_row.getCell(messageTimeIndex).toString());
                    String referenzzeitBeginn = convertDate(new_row.getCell(referencetimeBeginIndex).toString());
                    String referenzzeitEnde = convertDate(new_row.getCell(referencetimeEndIndex).toString());
                    String output = versandzeitpunkt + ";" + nachrichtenzeitpunkt + ";"
                    +new_row.getCell(nameIndex).toString() + ";" + referenzzeitBeginn + ";" + referenzzeitEnde
                    +";" + new_row.getCell(receivesenderIndex).toString().trim() + ";"
                    +new_row.getCell(receiveSearchTerm).toString() + ";"
                    +new_row.getCell(statusIndex).toString();
                    messageLog.setStringProperty("in loop",output)
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
