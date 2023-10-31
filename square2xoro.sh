#!/bin/sh
# input: Square website > Transactions > Export Items Detail CSV; Square website > Customers > Export Customers

cd ~/OneDrive\ -\ Jan\ and\ Jul/TWK\ -\ Gloria/VancouverBaby/
items="shipping-items-2023-10-28-2023-10-30.csv"
customers="customers-20231029.csv"
out="shipping-xoro-20231029.csv"
awk -F"," 'NR==FNR {first[$15]=$2; last[$15]=$3; email[$15]=$4; phone[$15]=$5; address1[$15]=$8; address2[$15]=$9; city[$15]=$10; state[$15]=$11; zip[$15]=$12; next} {print last[$22]","first[$22]","$5","$6","$12","address1[$22]","address2[$22]","city[$22]","state[$22]","zip[$22]","phone[$22]","email[$22]}' $customers $items > $out
