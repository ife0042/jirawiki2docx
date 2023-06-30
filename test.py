

from jirawiki2docx import JiraWiki2Docx

sample_jira_text = """Carried out an {color:#ff5630}analysis to {color}{color:#fff0b3}identify{color}{color:#ff5630} customers who{color} carried out POS/WEB (card) transactions during the period of +^*-_cash scarcity but subsequently_-*^+ stopped transacting on this *channel and map out the* *_transactional_* *behaviour* of _these_ *_customers_* _prior to and during_ cash scarcity.

*Definition of Terms*

* Cash scarcity period: November _*2022 to March 2023*_ (150 days)
** 6 months prior to cash scarc ^ity^ : May 2022 to October 2022 (180 days)
* Period after cash ~scarc~ity: April 2023 to May 2023 (60 days)
**** Period after cash scarcity: April 2023 to May 2023 (60 days)

# test number 1
#* test number 11
#* test number 11
# test number 2
## test number 2

*Key Highlights:*

* Customers transactional behaviour before, during and after cash scarcity

||Period||Transacted before, during {color:#ff5630}AND {color}after cash scarcity||Transacted during cash scarcity {color:#ff5630}BUT {color}stopped after cash scarcity||{color:#ff5630}STARTED {color}transacting during cash scarcity {color:#ff5630}AND {color}continued afterwards||
|* Distinct Card Holders
* dummy|3,674,893|{color:#ff5630}*1,310,454*{color}|602,235|

||* go
* be||
| |
| |


* *1.6 million (66%)* of the customers who completed card transactions during the cash scarcity period also carried out transactions prior to and after the cash scarcity period.
* *13% (1 million)* of the customers who carried out POS/WEB transactions during cash scarcity stopped card transactions at the end of the cash scarcity period
* Customers with POS/WEB transactions during cash scarcity but not after cash scarcity

||Period||Transacted prior to AND during cash scarcity BUT subsequently stopped||Transacted only during cash scarcity, NOT before or after cash scarcity||% of customers who transacted only during cash scarcity||
|Distinct Card Holders|{color:#36b37e}30,063 {color}|{color:#ff991f}20,391{color}|{color:#ff991f}41%{color}|

* *29% (3.03 of 11 million)* of these customers also carried out POS/WEB transactions in the 6 months prior to cash scarcity which means that they regularly transacted using their cards regardless of cash scarcity.-"""


sample_jira_text = """Carried out an {color:#ff5630}analysis to {color}{color:#fff0b3}identify{color}{color:#ff5630} customers who{color} carried out POS/WEB (card) transactions during the period of +^*-_cash scarcity but subsequently_-*^+ stopped transacting on this *channel and map out the* *_transactional_* *behaviour* of _these_ *_customers_* _prior to and during_ cash scarcity.

*Definition of Terms*"""

wk = JiraWiki2Docx(sample_jira_text)
wk.writeJira2Docx()
