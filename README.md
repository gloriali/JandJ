# Jan and Jul system management   
[README_RichmondPOS.md](./README_RichmondPOS.md): Richmond in-person sales instructions.   
[clover.R](./clover.R): manage Clover system - price, barcode, stock etc - update routinely; generate inventory transfer request from Surrey to Richmond - run at request and email Shikshit.    
[square2xoro.sh](./square2xoro.sh): prepare Square sales for xoro upload - run at request and upload to xoro.   
[square.R](./square.R): sync square inventory with xoro, price with woo - update routinely.    
[reward.R](./reward.R): add POS sales to yotpo rewards program system - update routinely.    
[customers.sh](./customers.sh): combine customers from Clover, Square, and Stripe.    
[POtracking.R](./POtracking.R): track China to global shipping PO. - discard, use python instead.   
[POtracking.ipynb](./POtracking.ipynb): track China to global shipping PO python Jupyter Notebook.  
[POtracking.py](./POtracking.py): track China to global shipping PO python script.  
