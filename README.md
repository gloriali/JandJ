# Jan and Jul system management   
[README_RichmondPOS.md](./README_RichmondPOS.md): Richmond in-person sales instructions.   
[square2xoro.sh](./square2xoro.sh): prepare Square sales for xoro upload - run at request and upload to xoro.   
[routine.R](./routine.R): combine all routine tasks in R here.    
[clover.R](./clover.R): manage Clover system - price, barcode, stock etc - update routinely. Generate inventory transfer request from Surrey to Richmond - run at request and email Shikshit.    
[square.R](./square.R): sync square inventory with xoro, price with woo - update routinely.    
[reward.R](./reward.R): add POS sales to yotpo rewards program system - update routinely.    
[customers.sh](./customers.sh): combine customers from Clover, Square, and Stripe.    
[Analysis.R](./Analysis.R): discontinued SKUs management - update routinely.    
[POtracking.py](./POtracking.py): track China to global shipping PO python script - run at request.  
[POtracking.ipynb](./POtracking.ipynb): track China to global shipping PO python Jupyter Notebook. - not updated, use python script in RStudio.    
[POtracking.R](./POtracking.R): track China to global shipping PO. - discard, use python instead.   
[FBArefill.R](./FBArefill.R): calculate FBA refill quantity for each SKU - run routinely and send to Shikshit.   
[XHS.R](./XHS.R): create and manage listing files - generate at request.  
[Netsuite.R](./Netsuite.R): manage Netsuite system   
