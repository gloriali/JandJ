# Jan and Jul system management   
[clover.R](./clover.R): manage Clover system - price, barcode, stock etc - update routinely; generate inventory request from Surrey to Richmond - run at request and email Shikshit.    
[customers.sh](./customers.sh): combine customers from Clover, Square, and Stripe.    
[square2xoro.sh](./square2xoro.sh): prepare Square sales for xoro upload - run at request and email Garman.   
[square.R](./square.R): sync square inventory with xoro, price with woo - update routinely.    
[reward.R](./reward.R): add POS sales to yotpo rewards program system - update routinely.    
[shippingTracking.R](./shippingTracking.R): track China to global shipping PO.   
