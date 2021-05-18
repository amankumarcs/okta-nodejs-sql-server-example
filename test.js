const { exist } = require("joi");
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('CBR');

var data = 
[{
    
    "Sr No" : "1",
    "Bank" : "BOI",
    "Feeder Branch" : "BOI UDUMALPETTAI",
    "ATM ID" : "CCB8061",
    "LOCATION" : "NO:35,MADURAI ROAD, OPP TO NEW POLICE STATION, DHARAPURAM -638656",
    "Date" : "2021-03-01 00:00:00",
    "Eod or Loading Time" : "00:00",
    "Status of Loading" : "NO EOD NO LOADING",
    "Last Transaction No" : "0",
    "CRA" : "SECURE VALUE",
    "Indent No" : "0",
    "Bank Ref No" : "0",
    "ATM counter details opening balance 100" : "0",
    "ATM counter details opening balance 200" : "0",
    "ATM counter details opening balance 500" : "1248500",
    "ATM counter details opening balance 1000" : "0",
    "ATM counter details opening balance 2000" : "0",
    "ATM counter details opening balance Total" : "1248500",
    "ATM counter details dispense counter 100" : "0",
    "ATM counter details dispense counter 200" : "0",
    "ATM counter details dispense counter 500" : "0",
    "ATM counter details dispense counter 1000" : "0",
    "ATM counter details dispense counter 2000" : "0",
    "ATM counter details dispense counter Total" : "0",
    "ATM counter details divert count 100" : "0",
    "ATM counter details divert count 200" : "0",
    "ATM counter details divert count 500" : "0",
    "ATM counter details divert count 1000" : "0",
    "ATM counter details divert count 2000" : "0",
    "ATM counter details divert count Total" : "0",
    "ATM counter details remaining counter 100" : "0",
    "ATM counter details remaining counter 200" : "0",
    "ATM counter details remaining counter 500" : "1248500",
    "ATM counter details remaining counter 1000" : "0",
    "ATM counter details remaining counter 2000" : "0",
    "ATM counter details remaining counter Total" : "1248500",
    "ATM  Details Physical cash from Cassettes 100" : "0",
    "ATM  Details Physical cash from Cassettes 200" : "0",
    "ATM  Details Physical cash from Cassettes 500" : "1248500",
    "ATM  Details Physical cash from Cassettes 1000" : "0",
    "ATM  Details Physical cash from Cassettes 2000" : "0",
    "ATM  Details Physical cash from Cassettes Total" : "1248500",
    "ATM  Details Physical total Cash from Purge Bin 100" : "0",
    "ATM  Details Physical total Cash from Purge Bin 200" : "0",
    "ATM  Details Physical total Cash from Purge Bin 500" : "0",
    "ATM  Details Physical total Cash from Purge Bin 1000" : "0",
    "ATM  Details Physical total Cash from Purge Bin 2000" : "0",
    "ATM  Details Physical total Cash from Purge Bin Total" : "0",
    "ATM  Details Physical total Remaining Cash 100" : "0",
    "ATM  Details Physical total Remaining Cash 200" : "0",
    "ATM  Details Physical total Remaining Cash 500" : "1248500",
    "ATM  Details Physical total Remaining Cash 1000" : "0",
    "ATM  Details Physical total Remaining Cash 2000" : "0",
    "ATM  Details Physical total Remaining Cash Total" : "1248500",
    "ATM Return Cash to Vault or Bank cash from Cassettes 100" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 200" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 500" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 1000" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 2000" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes Total" : "0",
    "ATM Return Cash to Vault or Bank seal no 100" : "0",
    "ATM Return Cash to Vault or Bank seal no 200" : "0",
    "ATM Return Cash to Vault or Bank seal no 500" : "0",
    "ATM Return Cash to Vault or Bank seal no 1000" : "0",
    "ATM Return Cash to Vault or Bank seal no 2000" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 100" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 200" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 500" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 1000" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 2000" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin Total" : "0",
    "ATM Repl Details Amount Replenished 100" : "0",
    "ATM Repl Details Amount Replenished 200" : "0",
    "ATM Repl Details Amount Replenished 500" : "0",
    "ATM Repl Details Amount Replenished 1000" : "0",
    "ATM Repl Details Amount Replenished 2000" : "0",
    "ATM Repl Details Amount Replenished Total" : "0",
    "ATM Repl Details seal no 100" : "0",
    "ATM Repl Details seal no 200" : "0",
    "ATM Repl Details seal no 500" : "0",
    "ATM Repl Details seal no 1000" : "0",
    "ATM Repl Details seal no 2000" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 100" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 200" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 500" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 1000" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 2000" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns Total" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM 100" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM 200" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM 500" : "1248500",
    "ATM Returns or Closing Balance Closing Balance in ATM 1000" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM 2000" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM Total" : "1248500",
    "Switch Counter Opening Balance 100" : "0",
    "Switch Counter Opening Balance 200" : "0",
    "Switch Counter Opening Balance 500" : "1248500",
    "Switch Counter Opening Balance 1000" : "0",
    "Switch Counter Opening Balance 2000" : "0",
    "Switch Counter Opening Balance Total" : "1248500",
    "Switch Counter Dispense as per Switch 100" : "0",
    "Switch Counter Dispense as per Switch 200" : "0",
    "Switch Counter Dispense as per Switch 500" : "0",
    "Switch Counter Dispense as per Switch 1000" : "0",
    "Switch Counter Dispense as per Switch 2000" : "0",
    "Switch Counter Dispense as per Switch Total" : "0",
    "Switch Counter Loading 100" : "0",
    "Switch Counter Loading 200" : "0",
    "Switch Counter Loading 500" : "0",
    "Switch Counter Loading 1000" : "0",
    "Switch Counter Loading 2000" : "0",
    "Switch Counter Loading Total" : "0",
    "Switch Counter Admin increase 100" : "0",
    "Switch Counter Admin increase 200" : "0",
    "Switch Counter Admin increase 500" : "0",
    "Switch Counter Admin increase 1000" : "0",
    "Switch Counter Admin increase 2000" : "0",
    "Switch Counter Admin increase Total" : "0",
    "Switch Counter Admin decrease 100" : "0",
    "Switch Counter Admin decrease 200" : "0",
    "Switch Counter Admin decrease 500" : "0",
    "Switch Counter Admin decrease 1000" : "0",
    "Switch Counter Admin decrease 2000" : "0",
    "Switch Counter Admin decrease Total" : "0",
    "Switch Counter Admin closing 100" : "0",
    "Switch Counter Admin closing 200" : "0",
    "Switch Counter Admin closing 500" : "1248500",
    "Switch Counter Admin closing 1000" : "0",
    "Switch Counter Admin closing 2000" : "0",
    "Switch Counter Admin closing Total" : "1248500",
    "Physical Difference Overage 100" : "0",
    "Physical Difference Overage 200" : "0",
    "Physical Difference Overage 500" : "0",
    "Physical Difference Overage 1000" : "0",
    "Physical Difference Overage 2000" : "0",
    "Physical Difference Overage Total" : "0",
    "Physical Difference Shortage 100" : "0",
    "Physical Difference Shortage 200" : "0",
    "Physical Difference Shortage 500" : "0",
    "Physical Difference Shortage 1000" : "0",
    "Physical Difference Shortage 2000" : "0",
    "Physical Difference Shortage Total" : "0",
    "Status one" : "OK",
    "Status two" : "OK",
    "Remarks" : "NO SCHEDULE",
    "ATM Unfit Currency 100" : "0",
    "ATM Unfit Currency 200" : "0",
    "ATM Unfit Currency 500" : "0",
    "ATM Unfit Currency 1000" : "0",
    "ATM Unfit Currency 2000" : "0",
    "ATM Unfit Currency Total" : "0",
    "id" : "test1234",
  
    "processName" : "ate",
    "clientName" : "test",
    "atmiddeno" : "CCB8061undefined",
    "atmiddenocount" : 1
},
{  
    "Sr No" : "2",
    "Bank" : "BOI",
    "Feeder Branch" : "CBD BELAPUR",
    "ATM ID" : "CRG8015",
    "LOCATION" : "AT POST JAMBHIVALI, TALUKA KARJAT,DIST-RAIGAD(OFF SITE)",
    "Date" : "2021-03-01 00:00:00",
    "Eod or Loading Time" : "14:33",
    "Status of Loading" : "LOADING",
    "Last Transaction No" : "0",
    "CRA" : "SECURE VALUE",
    "Indent No" : "0",
    "Bank Ref No" : "0",
    "ATM counter details opening balance 100" : "124200",
    "ATM counter details opening balance 200" : "0",
    "ATM counter details opening balance 500" : "496000",
    "ATM counter details opening balance 1000" : "0",
    "ATM counter details opening balance 2000" : "0",
    "ATM counter details opening balance Total" : "620200",
    "ATM counter details dispense counter 100" : "5300",
    "ATM counter details dispense counter 200" : "0",
    "ATM counter details dispense counter 500" : "197000",
    "ATM counter details dispense counter 1000" : "0",
    "ATM counter details dispense counter 2000" : "0",
    "ATM counter details dispense counter Total" : "202300",
    "ATM counter details divert count 100" : "0",
    "ATM counter details divert count 200" : "0",
    "ATM counter details divert count 500" : "0",
    "ATM counter details divert count 1000" : "0",
    "ATM counter details divert count 2000" : "0",
    "ATM counter details divert count Total" : "0",
    "ATM counter details remaining counter 100" : "118900",
    "ATM counter details remaining counter 200" : "0",
    "ATM counter details remaining counter 500" : "299000",
    "ATM counter details remaining counter 1000" : "0",
    "ATM counter details remaining counter 2000" : "0",
    "ATM counter details remaining counter Total" : "417900",
    "ATM  Details Physical cash from Cassettes 100" : "118900",
    "ATM  Details Physical cash from Cassettes 200" : "0",
    "ATM  Details Physical cash from Cassettes 500" : "299000",
    "ATM  Details Physical cash from Cassettes 1000" : "0",
    "ATM  Details Physical cash from Cassettes 2000" : "0",
    "ATM  Details Physical cash from Cassettes Total" : "417900",
    "ATM  Details Physical total Cash from Purge Bin 100" : "0",
    "ATM  Details Physical total Cash from Purge Bin 200" : "0",
    "ATM  Details Physical total Cash from Purge Bin 500" : "0",
    "ATM  Details Physical total Cash from Purge Bin 1000" : "0",
    "ATM  Details Physical total Cash from Purge Bin 2000" : "0",
    "ATM  Details Physical total Cash from Purge Bin Total" : "0",
    "ATM  Details Physical total Remaining Cash 100" : "118900",
    "ATM  Details Physical total Remaining Cash 200" : "0",
    "ATM  Details Physical total Remaining Cash 500" : "299000",
    "ATM  Details Physical total Remaining Cash 1000" : "0",
    "ATM  Details Physical total Remaining Cash 2000" : "0",
    "ATM  Details Physical total Remaining Cash Total" : "417900",
    "ATM Return Cash to Vault or Bank cash from Cassettes 100" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 200" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 500" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 1000" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 2000" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes Total" : "0",
    "ATM Return Cash to Vault or Bank seal no 100" : "0",
    "ATM Return Cash to Vault or Bank seal no 200" : "0",
    "ATM Return Cash to Vault or Bank seal no 500" : "0",
    "ATM Return Cash to Vault or Bank seal no 1000" : "0",
    "ATM Return Cash to Vault or Bank seal no 2000" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 100" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 200" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 500" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 1000" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 2000" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin Total" : "0",
    "ATM Repl Details Amount Replenished 100" : "100000",
    "ATM Repl Details Amount Replenished 200" : "0",
    "ATM Repl Details Amount Replenished 500" : "1000000",
    "ATM Repl Details Amount Replenished 1000" : "0",
    "ATM Repl Details Amount Replenished 2000" : "0",
    "ATM Repl Details Amount Replenished Total" : "1100000",
    "ATM Repl Details seal no 100" : "0",
    "ATM Repl Details seal no 200" : "0",
    "ATM Repl Details seal no 500" : "0",
    "ATM Repl Details seal no 1000" : "0",
    "ATM Repl Details seal no 2000" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 100" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 200" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 500" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 1000" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 2000" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns Total" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM 100" : "218900",
    "ATM Returns or Closing Balance Closing Balance in ATM 200" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM 500" : "1299000",
    "ATM Returns or Closing Balance Closing Balance in ATM 1000" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM 2000" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM Total" : "1517900",
    "Switch Counter Opening Balance 100" : "124200",
    "Switch Counter Opening Balance 200" : "0",
    "Switch Counter Opening Balance 500" : "496000",
    "Switch Counter Opening Balance 1000" : "0",
    "Switch Counter Opening Balance 2000" : "0",
    "Switch Counter Opening Balance Total" : "620200",
    "Switch Counter Dispense as per Switch 100" : "5300",
    "Switch Counter Dispense as per Switch 200" : "0",
    "Switch Counter Dispense as per Switch 500" : "197000",
    "Switch Counter Dispense as per Switch 1000" : "0",
    "Switch Counter Dispense as per Switch 2000" : "0",
    "Switch Counter Dispense as per Switch Total" : "202300",
    "Switch Counter Loading 100" : "100000",
    "Switch Counter Loading 200" : "0",
    "Switch Counter Loading 500" : "1000000",
    "Switch Counter Loading 1000" : "0",
    "Switch Counter Loading 2000" : "0",
    "Switch Counter Loading Total" : "1100000",
    "Switch Counter Admin increase 100" : "0",
    "Switch Counter Admin increase 200" : "0",
    "Switch Counter Admin increase 500" : "0",
    "Switch Counter Admin increase 1000" : "0",
    "Switch Counter Admin increase 2000" : "0",
    "Switch Counter Admin increase Total" : "0",
    "Switch Counter Admin decrease 100" : "0",
    "Switch Counter Admin decrease 200" : "0",
    "Switch Counter Admin decrease 500" : "0",
    "Switch Counter Admin decrease 1000" : "0",
    "Switch Counter Admin decrease 2000" : "0",
    "Switch Counter Admin decrease Total" : "0",
    "Switch Counter Admin closing 100" : "218900",
    "Switch Counter Admin closing 200" : "0",
    "Switch Counter Admin closing 500" : "1299000",
    "Switch Counter Admin closing 1000" : "0",
    "Switch Counter Admin closing 2000" : "0",
    "Switch Counter Admin closing Total" : "1517900",
    "Physical Difference Overage 100" : "0",
    "Physical Difference Overage 200" : "0",
    "Physical Difference Overage 500" : "0",
    "Physical Difference Overage 1000" : "0",
    "Physical Difference Overage 2000" : "0",
    "Physical Difference Overage Total" : "0",
    "Physical Difference Shortage 100" : "0",
    "Physical Difference Shortage 200" : "0",
    "Physical Difference Shortage 500" : "0",
    "Physical Difference Shortage 1000" : "0",
    "Physical Difference Shortage 2000" : "0",
    "Physical Difference Shortage Total" : "0",
    "Status one" : "OK",
    "Status two" : "OK",
    "Remarks" : "Loading done",
    "ATM Unfit Currency 100" : "0",
    "ATM Unfit Currency 200" : "0",
    "ATM Unfit Currency 500" : "0",
    "ATM Unfit Currency 1000" : "0",
    "ATM Unfit Currency 2000" : "0",
    "ATM Unfit Currency Total" : "0",
    "id" : "test1234",
    
    "processName" : "ate",
    "clientName" : "test",
    "atmiddeno" : "CRG8015undefined",
    "atmiddenocount" : 1
},
{
   
    "Sr No" : "3",
    "Bank" : "BOI",
    "Feeder Branch" : "BHOPAL",
    "ATM ID" : "CBO8052",
    "LOCATION" : "SHOP NO. 63 PRIYANKA LAJ MANDIDEEP BHOPAL MP (462046)",
    "Date" : "2021-03-01 00:00:00",
    "Eod or Loading Time" : "00:00",
    "Status of Loading" : "NO EOD NO LOADING",
    "Last Transaction No" : "0",
    "CRA" : "SECURE VALUE",
    "Indent No" : "0",
    "Bank Ref No" : "0",
    "ATM counter details opening balance 100" : "153300",
    "ATM counter details opening balance 200" : "267600",
    "ATM counter details opening balance 500" : "1571500",
    "ATM counter details opening balance 1000" : "0",
    "ATM counter details opening balance 2000" : "0",
    "ATM counter details opening balance Total" : "1992400",
    "ATM counter details dispense counter 100" : "0",
    "ATM counter details dispense counter 200" : "0",
    "ATM counter details dispense counter 500" : "0",
    "ATM counter details dispense counter 1000" : "0",
    "ATM counter details dispense counter 2000" : "0",
    "ATM counter details dispense counter Total" : "0",
    "ATM counter details divert count 100" : "0",
    "ATM counter details divert count 200" : "0",
    "ATM counter details divert count 500" : "0",
    "ATM counter details divert count 1000" : "0",
    "ATM counter details divert count 2000" : "0",
    "ATM counter details divert count Total" : "0",
    "ATM counter details remaining counter 100" : "153300",
    "ATM counter details remaining counter 200" : "267600",
    "ATM counter details remaining counter 500" : "1571500",
    "ATM counter details remaining counter 1000" : "0",
    "ATM counter details remaining counter 2000" : "0",
    "ATM counter details remaining counter Total" : "1992400",
    "ATM  Details Physical cash from Cassettes 100" : "153300",
    "ATM  Details Physical cash from Cassettes 200" : "267600",
    "ATM  Details Physical cash from Cassettes 500" : "1571500",
    "ATM  Details Physical cash from Cassettes 1000" : "0",
    "ATM  Details Physical cash from Cassettes 2000" : "0",
    "ATM  Details Physical cash from Cassettes Total" : "1992400",
    "ATM  Details Physical total Cash from Purge Bin 100" : "0",
    "ATM  Details Physical total Cash from Purge Bin 200" : "0",
    "ATM  Details Physical total Cash from Purge Bin 500" : "0",
    "ATM  Details Physical total Cash from Purge Bin 1000" : "0",
    "ATM  Details Physical total Cash from Purge Bin 2000" : "0",
    "ATM  Details Physical total Cash from Purge Bin Total" : "0",
    "ATM  Details Physical total Remaining Cash 100" : "153300",
    "ATM  Details Physical total Remaining Cash 200" : "267600",
    "ATM  Details Physical total Remaining Cash 500" : "1571500",
    "ATM  Details Physical total Remaining Cash 1000" : "0",
    "ATM  Details Physical total Remaining Cash 2000" : "0",
    "ATM  Details Physical total Remaining Cash Total" : "1992400",
    "ATM Return Cash to Vault or Bank cash from Cassettes 100" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 200" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 500" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 1000" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes 2000" : "0",
    "ATM Return Cash to Vault or Bank cash from Cassettes Total" : "0",
    "ATM Return Cash to Vault or Bank seal no 100" : "0",
    "ATM Return Cash to Vault or Bank seal no 200" : "0",
    "ATM Return Cash to Vault or Bank seal no 500" : "0",
    "ATM Return Cash to Vault or Bank seal no 1000" : "0",
    "ATM Return Cash to Vault or Bank seal no 2000" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 100" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 200" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 500" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 1000" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin 2000" : "0",
    "ATM Return Cash to Vault or Bank Total Cash from Purge Bin Total" : "0",
    "ATM Repl Details Amount Replenished 100" : "0",
    "ATM Repl Details Amount Replenished 200" : "0",
    "ATM Repl Details Amount Replenished 500" : "0",
    "ATM Repl Details Amount Replenished 1000" : "0",
    "ATM Repl Details Amount Replenished 2000" : "0",
    "ATM Repl Details Amount Replenished Total" : "0",
    "ATM Repl Details seal no 100" : "0",
    "ATM Repl Details seal no 200" : "0",
    "ATM Repl Details seal no 500" : "0",
    "ATM Repl Details seal no 1000" : "0",
    "ATM Repl Details seal no 2000" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 100" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 200" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 500" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 1000" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns 2000" : "0",
    "ATM Returns or Closing Balance Total ATM Cash Returns Total" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM 100" : "153300",
    "ATM Returns or Closing Balance Closing Balance in ATM 200" : "267600",
    "ATM Returns or Closing Balance Closing Balance in ATM 500" : "1571500",
    "ATM Returns or Closing Balance Closing Balance in ATM 1000" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM 2000" : "0",
    "ATM Returns or Closing Balance Closing Balance in ATM Total" : "1992400",
    "Switch Counter Opening Balance 100" : "153300",
    "Switch Counter Opening Balance 200" : "267600",
    "Switch Counter Opening Balance 500" : "1571500",
    "Switch Counter Opening Balance 1000" : "0",
    "Switch Counter Opening Balance 2000" : "0",
    "Switch Counter Opening Balance Total" : "1992400",
    "Switch Counter Dispense as per Switch 100" : "0",
    "Switch Counter Dispense as per Switch 200" : "0",
    "Switch Counter Dispense as per Switch 500" : "0",
    "Switch Counter Dispense as per Switch 1000" : "0",
    "Switch Counter Dispense as per Switch 2000" : "0",
    "Switch Counter Dispense as per Switch Total" : "0",
    "Switch Counter Loading 100" : "0",
    "Switch Counter Loading 200" : "0",
    "Switch Counter Loading 500" : "0",
    "Switch Counter Loading 1000" : "0",
    "Switch Counter Loading 2000" : "0",
    "Switch Counter Loading Total" : "0",
    "Switch Counter Admin increase 100" : "0",
    "Switch Counter Admin increase 200" : "0",
    "Switch Counter Admin increase 500" : "0",
    "Switch Counter Admin increase 1000" : "0",
    "Switch Counter Admin increase 2000" : "0",
    "Switch Counter Admin increase Total" : "0",
    "Switch Counter Admin decrease 100" : "0",
    "Switch Counter Admin decrease 200" : "0",
    "Switch Counter Admin decrease 500" : "0",
    "Switch Counter Admin decrease 1000" : "0",
    "Switch Counter Admin decrease 2000" : "0",
    "Switch Counter Admin decrease Total" : "0",
    "Switch Counter Admin closing 100" : "153300",
    "Switch Counter Admin closing 200" : "267600",
    "Switch Counter Admin closing 500" : "1571500",
    "Switch Counter Admin closing 1000" : "0",
    "Switch Counter Admin closing 2000" : "0",
    "Switch Counter Admin closing Total" : "1992400",
    "Physical Difference Overage 100" : "0",
    "Physical Difference Overage 200" : "0",
    "Physical Difference Overage 500" : "0",
    "Physical Difference Overage 1000" : "0",
    "Physical Difference Overage 2000" : "0",
    "Physical Difference Overage Total" : "0",
    "Physical Difference Shortage 100" : "0",
    "Physical Difference Shortage 200" : "0",
    "Physical Difference Shortage 500" : "0",
    "Physical Difference Shortage 1000" : "0",
    "Physical Difference Shortage 2000" : "0",
    "Physical Difference Shortage Total" : "0",
    "Status one" : "OK",
    "Status two" : "OK",
    "Remarks" : "NO EOD GIVEN BY MSP",
    "ATM Unfit Currency 100" : "0",
    "ATM Unfit Currency 200" : "0",
    "ATM Unfit Currency 500" : "0",
    "ATM Unfit Currency 1000" : "0",
    "ATM Unfit Currency 2000" : "0",
    "ATM Unfit Currency Total" : "0",
    "id" : "test1234",    
    "processName" : "ate",
    "clientName" : "test",
    "atmiddeno" : "CBO8052undefined",
    "atmiddenocount" : 1
}]

var finalData = [];


var keys = Object.keys(data[0]);
// console.log(keys);
var keysFor100 = [];
var keysFor200 = [];
var keysFor500 = [];
var keysFor1000 = [];
var keysFor2000 = [];
var keysForTotal = [];
keys.forEach(element => {
    
    if (element.endsWith("100")) {
        keysFor100.push(element);
    }
    if (element.endsWith("200")) {
        keysFor200.push(element);
    }
    if (element.endsWith("500")) {
        keysFor500.push(element);
    }
    if (element.endsWith("1000")) {
        keysFor1000.push(element);
    }
    if (element.endsWith("2000")) {
        keysFor2000.push(element);
    }
    if (element.endsWith("Total")) {
        keysForTotal.push(element);
    }
    
});

var testData = data[0];

let items = ["100","200","500","1000","2000","total"];
let key100 = keyFor100;
items.foreach((elem1)=>{
    data.forEach(element => {
    var data = {
        "Sr No" : element["Sr No"],
        "Bank" : element["Bank"] ,
        "Feeder Branch": element["Feeder Branch"],
        "ATM ID" : element["ATM ID"],
        "LOCATION" : element["LOCATION"],
        "Date" : element["Date"],
        "Eod or Loading Time" : element["Eod or Loading Time"],
        "Status of Loading" : element["Status of Loading"],
        "Last Transaction No" : element["Last Transaction No"],
        "CRA" : element["CRA"],
        "Indent No" : element["Indent No"],
        "Bank Ref No": element["Bank Ref No"],
    }

    (key + elem1).forEach(elementNew => {
      data[elementNew] = element[elementNew]
    });

    data100["Status one"] = element["Status one"];
    data100["Status two"] = element["Status two"];
    data100["Remarks"] = element["Remarks"];
    data100["processName"] = element["processName"];
    data100["clientName"] = element["clientName"];
    data100["id"] = element["id"];

    finalData.push(data)
});
});

// For 100
data.forEach(element => {
    var data100 = {
        "Sr No" : element["Sr No"],
        "Bank" : element["Bank"] ,
        "Feeder Branch": element["Feeder Branch"],
        "ATM ID" : element["ATM ID"],
        "LOCATION" : element["LOCATION"],
        "Date" : element["Date"],
        "Eod or Loading Time" : element["Eod or Loading Time"],
        "Status of Loading" : element["Status of Loading"],
        "Last Transaction No" : element["Last Transaction No"],
        "CRA" : element["CRA"],
        "Indent No" : element["Indent No"],
        "Bank Ref No": element["Bank Ref No"],
    }

    keysFor100.forEach(element100 => {
      data100[element100] = element[element100]
    });

    data100["Status one"] = element["Status one"];
    data100["Status two"] = element["Status two"];
    data100["Remarks"] = element["Remarks"];
    data100["processName"] = element["processName"];
    data100["clientName"] = element["clientName"];
    data100["id"] = element["id"];

    finalData.push(data100)
});
//  For 200
data.forEach(element200Data => {
    var data200 = {
        "Sr No" : element200Data["Sr No"],
        "Bank" : element200Data["Bank"] ,
        "Feeder Branch": element200Data["Feeder Branch"],
        "ATM ID" : element200Data["ATM ID"],
        "LOCATION" : element200Data["LOCATION"],
        "Date" : element200Data["Date"],
        "Eod or Loading Time" : element200Data["Eod or Loading Time"],
        "Status of Loading" : element200Data["Status of Loading"],
        "Last Transaction No" : element200Data["Last Transaction No"],
        "CRA" : element200Data["CRA"],
        "Indent No" : element200Data["Indent No"],
        "Bank Ref No": element200Data["Bank Ref No"],
    }
    
    keysFor200.forEach(element200 => {
        data200[element200] = element200Data[element200]
    });

    data200["Status one"] = element200Data["Status one"];
    data200["Status two"] = element200Data["Status two"];
    data200["Remarks"] = element200Data["Remarks"];
    data200["processName"] = element200Data["processName"];
    data200["clientName"] = element200Data["clientName"];
    data200["id"] = element200Data["id"];

    finalData.push(data200)
    
});


// For 500
data.forEach(element500Data => {
    var data500 = {
        "Sr No" : element500Data["Sr No"],
        "Bank" : element500Data["Bank"] ,
        "Feeder Branch": element500Data["Feeder Branch"],
        "ATM ID" : element500Data["ATM ID"],
        "LOCATION" : element500Data["LOCATION"],
        "Date" : element500Data["Date"],
        "Eod or Loading Time" : element500Data["Eod or Loading Time"],
        "Status of Loading" : element500Data["Status of Loading"],
        "Last Transaction No" : element500Data["Last Transaction No"],
        "CRA" : element500Data["CRA"],
        "Indent No" : element500Data["Indent No"],
        "Bank Ref No": element500Data["Bank Ref No"],
    }
    
    keysFor500.forEach(element500 => {
        data500[element500] = element500Data[element500]
    });

    data500["Status one"] = element500Data["Status one"];
    data500["Status two"] = element500Data["Status two"];
    data500["Remarks"] = element500Data["Remarks"];
    data500["processName"] = element500Data["processName"];
    data500["clientName"] = element500Data["clientName"];
    data500["id"] = element500Data["id"];

    finalData.push(data500)
    
});

// For 1000
data.forEach(element1000Data => {
    var data1000 = {
        "Sr No" : element1000Data["Sr No"],
        "Bank" : element1000Data["Bank"] ,
        "Feeder Branch": element1000Data["Feeder Branch"],
        "ATM ID" : element1000Data["ATM ID"],
        "LOCATION" : element1000Data["LOCATION"],
        "Date" : element1000Data["Date"],
        "Eod or Loading Time" : element1000Data["Eod or Loading Time"],
        "Status of Loading" : element1000Data["Status of Loading"],
        "Last Transaction No" : element1000Data["Last Transaction No"],
        "CRA" : element1000Data["CRA"],
        "Indent No" : element1000Data["Indent No"],
        "Bank Ref No": element1000Data["Bank Ref No"],
    }
    
    keysFor1000.forEach(element1000 => {
        data1000[element1000] = element1000Data[element1000]
    });

    data1000["Status one"] = element1000Data["Status one"];
    data1000["Status two"] = element1000Data["Status two"];
    data1000["Remarks"] = element1000Data["Remarks"];
    data1000["processName"] = element1000Data["processName"];
    data1000["clientName"] = element1000Data["clientName"];
    data1000["id"] = element1000Data["id"];

    finalData.push(data1000)
    
});

//  For 2000
data.forEach(element2000data => {
    var data2000 = {
        "Sr No" : element2000data["Sr No"],
        "Bank" : element2000data["Bank"] ,
        "Feeder Branch": element2000data["Feeder Branch"],
        "ATM ID" : element2000data["ATM ID"],
        "LOCATION" : element2000data["LOCATION"],
        "Date" : element2000data["Date"],
        "Eod or Loading Time" : element2000data["Eod or Loading Time"],
        "Status of Loading" : element2000data["Status of Loading"],
        "Last Transaction No" : element2000data["Last Transaction No"],
        "CRA" : element2000data["CRA"],
        "Indent No" : element2000data["Indent No"],
        "Bank Ref No": element2000data["Bank Ref No"],
    }
    
    keysFor2000.forEach(element2000 => {
        data2000[element2000] = element2000data[element2000]
    });

    data2000["Status one"] = element2000data["Status one"];
    data2000["Status two"] = element2000data["Status two"];
    data2000["Remarks"] = element2000data["Remarks"];
    data2000["processName"] = element2000data["processName"];
    data2000["clientName"] = element2000data["clientName"];
    data2000["id"] = element2000data["id"];

    finalData.push(data2000)
    
});


// For Total 

data.forEach(elementTotal => {
    var dataTotal = {
        "Sr No" : elementTotal["Sr No"],
        "Bank" : elementTotal["Bank"] ,
        "Feeder Branch": elementTotal["Feeder Branch"],
        "ATM ID" : elementTotal["ATM ID"],
        "LOCATION" : elementTotal["LOCATION"],
        "Date" : elementTotal["Date"],
        "Eod or Loading Time" : elementTotal["Eod or Loading Time"],
        "Status of Loading" : elementTotal["Status of Loading"],
        "Last Transaction No" : elementTotal["Last Transaction No"],
        "CRA" : elementTotal["CRA"],
        "Indent No" : elementTotal["Indent No"],
        "Bank Ref No": elementTotal["Bank Ref No"],
    }
    
    keysForTotal.forEach(elementTotal2 => {
        dataTotal[elementTotal2] = elementTotal[elementTotal2]
    });

    dataTotal["Status one"] = elementTotal["Status one"];
    dataTotal["Status two"] = elementTotal["Status two"];
    dataTotal["Remarks"] = elementTotal["Remarks"];
    dataTotal["processName"] = elementTotal["processName"];
    dataTotal["clientName"] = elementTotal["clientName"];
    dataTotal["id"] = elementTotal["id"];

    finalData.push(dataTotal)
    
});

console.log(finalData);

// data.forEach(record => {
//     let columnIndex = 1;
//     Object.keys(record).forEach(columnName => {
//         ws.cell(rowIndex, columnIndex++)
//             .string(record[columnName])
//     });
//     rowIndex++;
// });
//  wb.write('extractData.xlsx');
// console.log("File generated ");


    

