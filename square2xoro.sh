#!/bin/sh
# convert square sales report to xoro upload format
## Square website > Transactions > Export Items Detail CSV;
## Square website > Customers > Export Customers

cd ~/OneDrive\ -\ Jan\ and\ Jul/TWK\ -\ Gloria/Square/
items=$(ls items-$(date +%Y-%m-%d)*.csv); # Square website > Transactions > Export Items Detail CSV;
customers="customers.csv"; # Square website > Customers > Export Customers
out="square2xoro-"$(date +%Y%m%d)".csv";

echo $items
echo "Order Date,Order Number,Last Name,First Name,Main Phone,Main Email,Customer Note,Payment Terms,SKU,Unit Price,QTY,Shipping cost,Shipping Address,Shipping Address 2,Shipping City,Shipping State / Province,Shipping Country Code,Shipping ZipCode,PO No" > $out
awk -F"," 'NR==FNR {gsub("\047\\+", "", $5); first[$15]=$2; last[$15]=$3; email[$15]=$4; phone[$15]=$5; address1[$15]=$8; address2[$15]=$9; city[$15]=$10; state[$15]=$11; zip[$15]=$12; next} FNR>1{IGNORECASE=1; date=$1; gsub("-", "", $1); gsub(/[$,"]/, "", $12); if($6==""){price=""}else{price=int($12/$6*100)/100}; if($14!=id){id=$14; no++}; if(first[$22]=="" && last[$22]==""){l="Walk-in"; f=$20} else{l=last[$22]; f=first[$22]}; if($20=="Surrey"){$17="Walk in customer-Picked"}; if($0 ~ /R Pickup/ || address1[$22] ~ /R Pickup/ || $0 ~ /Richmond Pickup/ || address1[$22] ~ /Richmond Pickup/){a1="13900 Maycrest Way Unit 135"; a2=""; c="Richmond"; s="BC"; z="V6V 3E2"} else if($0 ~ /U Pickup/ || address1[$22] ~ /U Pickup/ || $0 ~ /UBC Pickup/ || address1[$22] ~ /UBC Pickup/){a1="6272 Logan Lane"; a2=""; c="Vancouver"; s="BC"; z="V6T 2K9"} else if($0 ~ /S Pickup/ || address1[$22] ~ /S Pickup/ || $0 ~ /Surrey Pickup/ || address1[$22] ~ /Surrey Pickup/ || $20=="Surrey" ){a1="3577 194 St"; a2="Unit 106"; c="Surrey"; s="BC"; z="V3Z 1A5"} else{a1=address1[$22]; a2=address2[$22]; c=city[$22]; s=state[$22]; z=zip[$22]}; printf("%s,POP%s-%02d", date, $1, no); print ","l","f","phone[$22]","email[$22]","$17",CIA,"$5","price","$6",,"a1","a2","c","s",CA,"z","$14}' $customers $items >> $out 
# upload to xoro > Utilities > Data Imports > Upload Sales Orders - Utility Template: Popup Shop Orders CA
