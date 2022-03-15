# Meniga Categorize Excel
This tool is used to enrich transactions from an Excel file. The Excel file should have a list of transactions in one of the sheets and first row contains the header names. The header names are resolved into transaction property names, where spaces and underscores are removed and the property made camel case. Following standard column names are possible:
* identifier
* text
* currency
* counterpartyAccountId
* counterpartyName
* TerminalId
* externalMerchantId
* merchantName
* countryCode
* city
* street
* postalCode
* region
* geoLocation
* maskedPan
* checkId
* purposeCode
* bankTransactionCode
* creditorId
* reference
* transactionDate
* bookingDate
* valueDate
* timestamp
* mcc
* amount
* amountInCurrency
* bookedAmount
* accountBalance
* isMerchant
* isOwnAccountTransfer
* isPending  

Variants of those names are also allowed such as "Booked Amount", "MCC", "Text", "counterparty_name", etc.

The Excel document can also contain columns with CategoryId, CategoryName and NormalizedText/CleanedText that contain expected values from the categorization process. If those columns exists (or variants of them), the categorization results will be compared to values of those columns and marked with green or red depending on if the results were as expected or not.