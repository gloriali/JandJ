#!/bin/sh
# input: Square website > Customers > Export Customers; Clover website > Customers > Download; Stripe website > Customers > Export -> convert csv to tab delimited text to avoid problems caused by "," inside column "Description"

cd ~/OneDrive\ -\ Jan\ and\ Jul/TWK\ -\ Gloria/Customers/
clover="Customers-20231030-clover.csv"
square="export-20231030-square.csv"
stripe="stripe_customers_20231030.txt"
out="Customers-20231030-all.csv"
echo -e "\"Source\",\"ID\",\"First Name\",\"Last Name\",\"Email\",\"Phone\",\"Address Line1\",\"Address Line2\",\"City\",\"State\",\"Country\",\"Postal Code\"" > $out
less $clover | awk -F"," 'NR>1{print "\"Clover\",\""$1"\",\""$2"\",\""$3"\",\""$5"\",\""$4"\",\""$6"\",\""$7"\",\""$9"\",\""$10"\",\""$12"\",\""$11"\""}' >> $out
less $square | awk -F"," 'NR>1{print "\"Square\",\""$15"\",\""$2"\",\""$3"\",\""$4"\",\""$5"\",\""$8"\",\""$9"\",\""$10"\",\""$11"\",\"CA\",\""$12"\""}' >> $out
less $stripe | awk -F"\\t" 'NR>1{print "\"Stripe\",\""$1"\",\""$4"\",\"\",\""$3"\",\"""\",\""$14"\",\""$15"\",\""$16"\",\""$17"\",\""$18"\",\""$19"\""}' >> $out


