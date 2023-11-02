#!/bin/sh

cd ~/OneDrive\ -\ Jan\ and\ Jul/TWK\ -\ Gloria/Square/
items="items-2023-10-30-2023-10-31.csv" # Square website > Transactions > Export Items Detail CSV;
customers="customers.csv" # Square website > Customers > Export Customers
out="square2xoro-"$(date +%Y%m%d)".csv"
echo -e "Order Date,Order Number,Last Name,First Name,Main Phone,Main Email,Customer Note,Payment Terms,SKU,Unit Price,QTY,Shipping cost,Shipping Address,Shipping Address 2,Shipping City,Shipping State / Province,Shipping Country Code,Shipping ZipCode" > $out
awk -F"," 'NR==FNR {gsub("\047\\+", "", $5); first[$15]=$2; last[$15]=$3; email[$15]=$4; phone[$15]=$5; address1[$15]=$8; address2[$15]=$9; city[$15]=$10; state[$15]=$11; zip[$15]=$12; next} FNR>1{IGNORECASE=1; if($6==""){price=""}else{price=$12/$6}; if($0 ~ /Richmond Pickup/){a1="13900 Maycrest Way Unit 135"; a2=""; c="Richmond"; s="BC"; z="V6V 3E2"} else if($0 ~ /UBC Pickup/){a1="6272 Logan Lane"; a2=""; c="Vancouver"; s="BC"; z="V6T 2K9"} else{a1=address1[$22]; a2=address2[$22]; c=city[$22]; s=state[$22]; z=zip[$22]}; print $1","$14","last[$22]","first[$22]","phone[$22]","email[$22]","$17",CIA,"$5","price","$6",,"a1","a2","c","s",CA,"z}' $customers $items >> $out 

 