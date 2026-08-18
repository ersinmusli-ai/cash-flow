// DATA.JS - SenFood Skopje Cash Flow 2026
// Генерирано автоматски од Excel: 0 фактури, 126 АЛДАКО записи, 186 ИМПЕРИЈАЛ записи

const EXCHANGE_RATE = 61.7;

const CASH_DATA = {
  rate: EXCHANGE_RATE,
  months: ["January","February","March","April","May","June","July","August","September","October","November","December"],
  rows: {
    invoiceAldaco:    [210760.34, 45278.06, 292136.39, 101002.16, 37802.51, 227039.04, 135300.26, 88982.06, 0, 0, 0, 0],
    invoiceImperijal: [82502.61, 112402.45, 102404.46, 134069.34, 64235.35, 129554.59, 132011.78, 89397.94, 0, 0, 0, 0],
    transferFrom2025: [150044.0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    other:            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    paidInvoices:     [311442.0, 90532.0, 351107.0, 197624.3, 36070.0, 245749.6, 264777.0, 169446.6, 0, 0, 0, 0],
    exchangeDiff:     [1009.54, 293.46, 1138.11, 640.6, 116.92, 796.6, 858.27, 549.26, 0, 0, 0, 0],
    cashFlowTransfer: [20000.0, 30000.0, 20000.0, 15000.0, 0, 20000.0, 10000.0, 0, 0, 0, 0, 0],
    spedision:        [509.0, 629.0, 502.0, 370.0, 756.0, 566.0, 307.0, 0, 0, 0, 0, 0],
    bankCharges:      [38.0, 124.0, 86.0, 25.0, 0, 0, 0, 0, 0, 0, 0, 0],
    custom:           [31477.0, 30312.0, 28011.0, 17872.0, 44712.0, 42947.0, 16843.0, 0, 0, 0, 0, 0],
    variousCosts:     [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    incomeDDV:        [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    expensesDDV:      [12368.0, 11801.0, 10936.0, 6890.0, 17509.0, 16744.0, 6577.0, 0, 0, 0, 0, 0],
  },
  rowLabels: {
    invoiceAldaco:    'AMOUNT OF INVOICE "ALDACO"',
    invoiceImperijal: 'AMOUNT OF INVOICE "IMPERIJAL TABAKO"',
    transferFrom2025: 'Transfer from 2025',
    other:            'Other',
    paidInvoices:     'Senkardes PAID INVOICES',
    exchangeDiff:     'Exchange difference',
    cashFlowTransfer: 'SenFood Cash flow for transfer',
    spedision:        'SPEDISION MKD & SKYSPED',
    bankCharges:      'BANK CHARGES',
    custom:           'CUSTOM',
    variousCosts:     'Various costs',
    incomeDDV:        'INCOME DDV',
    expensesDDV:      'EXPENSES DDV',
  }
};

const INPORT_DATA = [
  {
    "month": "January",
    "invoice": "SKK2026000000002 22.01.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00376-26",
    "amountInv": 61690.2,
    "transport": 1100.0,
    "spedision": 136.0,
    "ddv": 3585.0,
    "custom": 8928.0,
    "banka": 38.0,
    "various": 0.0,
    "skAmount": 79175.0,
    "payMonth": 3,
    "status": "unpaid"
  },
  {
    "month": "January",
    "invoice": "SKK2026000000004 23.01.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00377-26",
    "amountInv": 76359.3,
    "transport": 1100.0,
    "spedision": 134.0,
    "ddv": 4434.0,
    "custom": 11240.0,
    "banka": 45.0,
    "various": 0.0,
    "skAmount": 97734.0,
    "payMonth": 4,
    "status": "unpaid"
  },
  {
    "month": "January",
    "invoice": "SKK2026000000003 23.01.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00378-26",
    "amountInv": 38580.0,
    "transport": 1050.0,
    "spedision": 119.0,
    "ddv": 2277.0,
    "custom": 5924.0,
    "banka": 26.0,
    "various": 0.0,
    "skAmount": 50810.0,
    "payMonth": 4,
    "status": "unpaid"
  },
  {
    "month": "January",
    "invoice": "SKK2026000000006 30.01.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00379-26",
    "amountInv": 35022.4,
    "transport": 1050.0,
    "spedision": 120.0,
    "ddv": 2072.0,
    "custom": 5385.0,
    "banka": 25.0,
    "various": 0.0,
    "skAmount": 45984.0,
    "payMonth": 3,
    "status": "unpaid"
  },
  {
    "month": "February",
    "invoice": "SKK2026000000007 04.02.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00380-26",
    "amountInv": 40547.0,
    "transport": 1050.0,
    "spedision": 126.0,
    "ddv": 2395.0,
    "custom": 6214.0,
    "banka": 27.0,
    "various": 0.0,
    "skAmount": 53389.0,
    "payMonth": 3,
    "status": "unpaid"
  },
  {
    "month": "February",
    "invoice": "SKK2026000000014 12.02.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00381-26",
    "amountInv": 39611.0,
    "transport": 1050.0,
    "spedision": 120.0,
    "ddv": 2341.0,
    "custom": 6084.0,
    "banka": 24.0,
    "various": 0.0,
    "skAmount": 52127.0,
    "payMonth": 3,
    "status": "unpaid"
  },
  {
    "month": "February",
    "invoice": "SKK2026000000015 12.02.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00382-26",
    "amountInv": 51231.4,
    "transport": 1100.0,
    "spedision": 136.0,
    "ddv": 2992.0,
    "custom": 7409.0,
    "banka": 24.0,
    "various": 0.0,
    "skAmount": 66315.0,
    "payMonth": 3,
    "status": "unpaid"
  },
  {
    "month": "February",
    "invoice": "SKK2026000000019 13.02.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00383-26",
    "amountInv": 33591.0,
    "transport": 1050.0,
    "spedision": 128.0,
    "ddv": 1996.0,
    "custom": 5206.0,
    "banka": 24.0,
    "various": 0.0,
    "skAmount": 44230.0,
    "payMonth": 3,
    "status": "unpaid"
  },
  {
    "month": "February",
    "invoice": "SKK2026000000029 20.02.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00384-26",
    "amountInv": 35020.0,
    "transport": 1050.0,
    "spedision": 119.0,
    "ddv": 2077.0,
    "custom": 5399.0,
    "banka": 25.0,
    "various": 0.0,
    "skAmount": 46065.0,
    "payMonth": 5,
    "status": "unpaid"
  },
  {
    "month": "February",
    "invoice": "SKK2026000000034 27.02.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00386-26",
    "amountInv": 45245.0,
    "transport": 1100.0,
    "spedision": 128.0,
    "ddv": 2659.0,
    "custom": 6743.0,
    "banka": 30.0,
    "various": 0.0,
    "skAmount": 59023.0,
    "payMonth": 4,
    "status": "unpaid"
  },
  {
    "month": "February",
    "invoice": "SKK2026000000035 03.02.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00387-26",
    "amountInv": 59036.2,
    "transport": 1100.0,
    "spedision": 119.0,
    "ddv": 3459.0,
    "custom": 8872.0,
    "banka": 32.0,
    "various": 0.0,
    "skAmount": 75787.0,
    "payMonth": 6,
    "status": "unpaid"
  },
  {
    "month": "February",
    "invoice": "SKK2026000000038 06.03.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00388-26",
    "amountInv": 33140.0,
    "transport": 1050.0,
    "spedision": 119.0,
    "ddv": 1972.0,
    "custom": 5144.0,
    "banka": 24.0,
    "various": 0.0,
    "skAmount": 43597.0,
    "payMonth": 4,
    "status": "unpaid"
  },
  {
    "month": "February",
    "invoice": "SKK2026000000050 19.03.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00389-26",
    "amountInv": 48420.0,
    "transport": 1100.0,
    "spedision": 136.0,
    "ddv": 2846.0,
    "custom": 7252.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 62774.0,
    "payMonth": 7,
    "status": "unpaid"
  },
  {
    "month": "April",
    "invoice": "SKK2026000000069 17.04.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00390-26",
    "amountInv": 35233.0,
    "transport": 1050.0,
    "spedision": 119.0,
    "ddv": 2091.0,
    "custom": 5441.0,
    "banka": 25.0,
    "various": 0.0,
    "skAmount": 46324.0,
    "payMonth": 6,
    "status": "unpaid"
  },
  {
    "month": "April",
    "invoice": "SKK2026000000070 21.04.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00391-26",
    "amountInv": 34344.0,
    "transport": 1050.0,
    "spedision": 119.0,
    "ddv": 2041.0,
    "custom": 5321.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 45198.0,
    "payMonth": 6,
    "status": "unpaid"
  },
  {
    "month": "April",
    "invoice": "SKK2026000000073 27.04.2026",
    "date": "",
    "client": "Ipmerijal",
    "invNo": "00392-26",
    "amountInv": 46711.4,
    "transport": 1200.0,
    "spedision": 132.0,
    "ddv": 2758.0,
    "custom": 7110.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 60614.0,
    "payMonth": 7,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000080 04.05.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00394-26",
    "amountInv": 50751.4,
    "transport": 1200.0,
    "spedision": 136.0,
    "ddv": 2977.0,
    "custom": 7441.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 65479.0,
    "payMonth": 6,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000086 12.05.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00395-26",
    "amountInv": 60495.0,
    "transport": 1490.0,
    "spedision": 136.0,
    "ddv": 3559.0,
    "custom": 8981.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 78201.0,
    "payMonth": 6,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000089 13.05.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00396-26",
    "amountInv": 37760.0,
    "transport": 1250.0,
    "spedision": 110.0,
    "ddv": 2242.0,
    "custom": 5823.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 49654.0,
    "payMonth": 7,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000090 15.05.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00397-26",
    "amountInv": 83906.0,
    "transport": 1200.0,
    "spedision": 136.0,
    "ddv": 4892.0,
    "custom": 12460.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 106980.0,
    "payMonth": 8,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000091 18.05.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00398-26",
    "amountInv": 31563.2,
    "transport": 1150.0,
    "spedision": 119.0,
    "ddv": 1886.0,
    "custom": 4912.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 41513.0,
    "payMonth": 7,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000096 27.05.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00399-26",
    "amountInv": 32714.4,
    "transport": 1150.0,
    "spedision": 119.0,
    "ddv": 1953.0,
    "custom": 5095.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 43045.0,
    "payMonth": 7,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000106 08.06.2026",
    "date": "",
    "client": "Aldaco",
    "invNo": "00400-26",
    "amountInv": 60558.0,
    "transport": 1200.0,
    "spedision": 119.0,
    "ddv": 3550.0,
    "custom": 9048.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 77930.0,
    "payMonth": 7,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000113 12.06.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00401-26",
    "amountInv": 42380.0,
    "transport": 1300.0,
    "spedision": 120.0,
    "ddv": 2517.0,
    "custom": 6521.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 55719.0,
    "payMonth": 0,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000116 24.06.2026",
    "date": "",
    "client": "Aldako",
    "invNo": "00402-26",
    "amountInv": 34450.0,
    "transport": 1150.0,
    "spedision": 120.0,
    "ddv": 2053.0,
    "custom": 5354.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 45328.0,
    "payMonth": 8,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000117 25.06.2026",
    "date": "",
    "client": "Aldako",
    "invNo": "00403-26",
    "amountInv": 47540.6,
    "transport": 1200.0,
    "spedision": 134.0,
    "ddv": 2801.0,
    "custom": 7130.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 61494.0,
    "payMonth": 8,
    "status": "unpaid"
  },
  {
    "month": "May",
    "invoice": "SKK2026000000118 26.06.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00404-26",
    "amountInv": 100045.4,
    "transport": 1200.0,
    "spedision": 73.0,
    "ddv": 5823.0,
    "custom": 14894.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 127187.0,
    "payMonth": 0,
    "status": "unpaid"
  },
  {
    "month": "July",
    "invoice": "SKK2026000000121 01.07.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00405-26",
    "amountInv": 79482.0,
    "transport": 1200.0,
    "spedision": 197.0,
    "ddv": 4632.0,
    "custom": 11799.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 102108.0,
    "payMonth": 0,
    "status": "unpaid"
  },
  {
    "month": "July",
    "invoice": "SKK2026000000133 30.07.2026",
    "date": "",
    "client": "Imperijal",
    "invNo": "00407-26",
    "amountInv": 32650.9,
    "transport": 1150.0,
    "spedision": 110.0,
    "ddv": 1945.0,
    "custom": 5044.0,
    "banka": 0.0,
    "various": 0.0,
    "skAmount": 42941.0,
    "payMonth": 0,
    "status": "unpaid"
  }
];

const CARD_ALDACO = [
  {
    "date": "17.01.2025",
    "desc": "Ф-ра 00323-25",
    "dolgува": 2777730.0,
    "pobaruva": 0,
    "valuta": "03.03.2025",
    "status": "paid"
  },
  {
    "date": "17.01.2025",
    "desc": "Аванс по Ф-ра 00323-25",
    "dolgува": 0,
    "pobaruva": 457368.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.01.2025",
    "desc": "Уплата по Ф-ра 00319-24",
    "dolgува": 0,
    "pobaruva": 0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.01.2025",
    "desc": "Уплата",
    "dolgува": 0,
    "pobaruva": 0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.01.2025",
    "desc": "Уплата по Ф-ра 00322-24",
    "dolgува": 0,
    "pobaruva": 0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.02.2025",
    "desc": "Ф-ра 00325-25",
    "dolgува": 5854776.0,
    "pobaruva": 0,
    "valuta": "22.03.2025",
    "status": "paid"
  },
  {
    "date": "05.02.2025",
    "desc": "Аванс по Ф-ра 00325-25",
    "dolgува": 0,
    "pobaruva": 944430.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.02.2025",
    "desc": "Ф-ра 00327-25",
    "dolgува": 2512125.0,
    "pobaruva": 0,
    "valuta": "22.03.2025",
    "status": "paid"
  },
  {
    "date": "05.02.2025",
    "desc": "Аванс по Ф-ра 00327-25",
    "dolgува": 0,
    "pobaruva": 414569.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.02.2025",
    "desc": "Ф-ра 00329-25",
    "dolgува": 5163990.0,
    "pobaruva": 0,
    "valuta": "11.04.2025",
    "status": "paid"
  },
  {
    "date": "25.02.2025",
    "desc": "Аванс по Ф-ра 00329-25",
    "dolgува": 0,
    "pobaruva": 834033.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.02.2025",
    "desc": "Уплата по Ф-ра 00323-25",
    "dolgува": 0,
    "pobaruva": 2320362.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.02.2025",
    "desc": "Уплата по Ф-ра 00325-25",
    "dolgува": 0,
    "pobaruva": 4910346.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.02.2025",
    "desc": "Уплата по Ф-ра 00327-25",
    "dolgува": 0,
    "pobaruva": 2097556.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "03.03.2025",
    "desc": "Ф-ра 00330-25",
    "dolgува": 2472934.0,
    "pobaruva": 0,
    "valuta": "17.04.2025",
    "status": "paid"
  },
  {
    "date": "03.03.2025",
    "desc": "Аванс по Ф-ра 00330-25",
    "dolgува": 0,
    "pobaruva": 408662.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.03.2025",
    "desc": "Ф-ра 00334-25",
    "dolgува": 2646930.0,
    "pobaruva": 0,
    "valuta": "28.04.2025",
    "status": "paid"
  },
  {
    "date": "14.03.2025",
    "desc": "Аванс по Ф-ра 00334-25",
    "dolgува": 0,
    "pobaruva": 434928.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.03.2025",
    "desc": "Ф-ра 00335-25",
    "dolgува": 3411396.0,
    "pobaruva": 0,
    "valuta": "28.04.2025",
    "status": "paid"
  },
  {
    "date": "14.03.2025",
    "desc": "Аванс по Ф-ра 00335-25",
    "dolgува": 0,
    "pobaruva": 551453.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "02.04.2025",
    "desc": "Ф-ра 00337-25",
    "dolgува": 2526156.0,
    "pobaruva": 0,
    "valuta": "17.05.2025",
    "status": "paid"
  },
  {
    "date": "02.04.2025",
    "desc": "Аванс по Ф-ра 00337-25",
    "dolgува": 0,
    "pobaruva": 417628.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.04.2025",
    "desc": "Ф-ра 00338-25",
    "dolgува": 4345332.0,
    "pobaruva": 0,
    "valuta": "22.05.2025",
    "status": "paid"
  },
  {
    "date": "07.04.2025",
    "desc": "Аванс по Ф-ра 00338-25",
    "dolgува": 0,
    "pobaruva": 703607.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "11.04.2025",
    "desc": "Уплата по Ф-ра 00329-25",
    "dolgува": 0,
    "pobaruva": 4329957.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.04.2025",
    "desc": "Уплата по Ф-ра 00330-25",
    "dolgува": 0,
    "pobaruva": 2064272.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "28.04.2025",
    "desc": "Уплата по Ф-ра 00334-25",
    "dolgува": 0,
    "pobaruva": 2212002.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "28.04.2025",
    "desc": "Уплата по Ф-ра 00335-25",
    "dolgува": 0,
    "pobaruva": 2859943.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.05.2025",
    "desc": "Ф-ра 00343-25",
    "dolgува": 2757762.0,
    "pobaruva": 0,
    "valuta": "28.06.2025",
    "status": "paid"
  },
  {
    "date": "14.05.2025",
    "desc": "Аванс по Ф-ра 00343-25",
    "dolgува": 0,
    "pobaruva": 449679.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "16.05.2025",
    "desc": "Уплата по Ф-ра 00337-25",
    "dolgува": 0,
    "pobaruva": 2108528.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.05.2025",
    "desc": "Уплата по Ф-ра 00338-25",
    "dolgува": 0,
    "pobaruva": 3641725.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "11.06.2025",
    "desc": "Ф-ра 00345-25",
    "dolgува": 2627708.0,
    "pobaruva": 0,
    "valuta": "26.07.2025",
    "status": "paid"
  },
  {
    "date": "11.06.2025",
    "desc": "Аванс по Ф-ра 00345-25",
    "dolgува": 0,
    "pobaruva": 428740.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.06.2025",
    "desc": "Ф-ра 00346-25",
    "dolgува": 2228580.0,
    "pobaruva": 0,
    "valuta": "04.08.2025",
    "status": "paid"
  },
  {
    "date": "20.06.2025",
    "desc": "Аванс по Ф-ра 00346-25",
    "dolgува": 0,
    "pobaruva": 371402.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "27.06.2025",
    "desc": "Уплата по Ф-ра 00343-25",
    "dolgува": 0,
    "pobaruva": 2308083.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "16.07.2025",
    "desc": "Ф-ра 00349-25",
    "dolgува": 3930010.0,
    "pobaruva": 0,
    "valuta": "30.08.2025",
    "status": "paid"
  },
  {
    "date": "16.07.2025",
    "desc": "Аванс по Ф-ра 00349-25",
    "dolgува": 0,
    "pobaruva": 638073.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.07.2025",
    "desc": "Уплата по Ф-ра 00345-25",
    "dolgува": 0,
    "pobaruva": 2198968.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.07.2025",
    "desc": "Уплата по Ф-ра 00346-25",
    "dolgува": 0,
    "pobaruva": 1857178.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "13.08.2025",
    "desc": "Ф-ра 00352-25",
    "dolgува": 2774153.0,
    "pobaruva": 0,
    "valuta": "27.09.2025",
    "status": "paid"
  },
  {
    "date": "13.08.2025",
    "desc": "Аванс по Ф-ра 00352-25",
    "dolgува": 0,
    "pobaruva": 467272.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "15.08.2025",
    "desc": "Ф-ра 00353-25",
    "dolgува": 3265951.0,
    "pobaruva": 0,
    "valuta": "29.09.2025",
    "status": "paid"
  },
  {
    "date": "15.08.2025",
    "desc": "Аванс по Ф-ра 00353-25",
    "dolgува": 0,
    "pobaruva": 533982.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "29.08.2025",
    "desc": "Уплата по Ф-ра 00349-25",
    "dolgува": 0,
    "pobaruva": 3291937.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.09.2025",
    "desc": "Ф-ра 00355-25",
    "dolgува": 2657592.0,
    "pobaruva": 0,
    "valuta": "25.10.2025",
    "status": "paid"
  },
  {
    "date": "10.09.2025",
    "desc": "Аванс по Ф-ра 00355-25",
    "dolgува": 0,
    "pobaruva": 448030.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "18.09.2025",
    "desc": "Ф-ра 00359-25",
    "dolgува": 3329440.0,
    "pobaruva": 0,
    "valuta": "02.11.2025",
    "status": "paid"
  },
  {
    "date": "18.09.2025",
    "desc": "Аванс по Ф-ра 00359-25",
    "dolgува": 0,
    "pobaruva": 539918.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "29.09.2025",
    "desc": "Уплата по Ф-ра 00352-25",
    "dolgува": 0,
    "pobaruva": 2306881.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "29.09.2025",
    "desc": "Уплата по Ф-ра 00353-25",
    "dolgува": 0,
    "pobaruva": 2731969.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "06.10.2025",
    "desc": "Ф-ра 00362-25",
    "dolgува": 3830214.0,
    "pobaruva": 0,
    "valuta": "20.11.2025",
    "status": "paid"
  },
  {
    "date": "06.10.2025",
    "desc": "Аванс по Ф-ра 00362-25",
    "dolgува": 0,
    "pobaruva": 614121.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "09.10.2025",
    "desc": "Ф-ра 00364-25",
    "dolgува": 2756736.0,
    "pobaruva": 0,
    "valuta": "23.11.2025",
    "status": "paid"
  },
  {
    "date": "09.10.2025",
    "desc": "Аванс по Ф-ра 00364-25",
    "dolgува": 0,
    "pobaruva": 462492.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "24.10.2025",
    "desc": "Уплата по Ф-ра 00355-25",
    "dolgува": 0,
    "pobaruva": 2209562.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "31.10.2025",
    "desc": "Уплата по Ф-ра 00359-25",
    "dolgува": 0,
    "pobaruva": 2789522.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "12.11.2025",
    "desc": "Ф-ра 00367-25",
    "dolgува": 4000578.0,
    "pobaruva": 0,
    "valuta": "27.12.2025",
    "status": "paid"
  },
  {
    "date": "12.11.2025",
    "desc": "Аванс по Ф-ра 00367-25",
    "dolgува": 0,
    "pobaruva": 649963.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.11.2025",
    "desc": "Ф-ра 00368-25",
    "dolgува": 3359768.0,
    "pobaruva": 0,
    "valuta": "04.01.2026",
    "status": "paid"
  },
  {
    "date": "20.11.2025",
    "desc": "Аванс по Ф-ра 00368-25",
    "dolgува": 0,
    "pobaruva": 557805.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.11.2025",
    "desc": "Уплата по Ф-ра 00362-25",
    "dolgува": 0,
    "pobaruva": 3216093.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.11.2025",
    "desc": "Уплата по Ф-ра 00364-25",
    "dolgува": 0,
    "pobaruva": 2294244.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.12.2025",
    "desc": "Ф-ра 00370-25",
    "dolgува": 2953800.0,
    "pobaruva": 0,
    "valuta": "24.01.2026",
    "status": "paid"
  },
  {
    "date": "10.12.2025",
    "desc": "Аванс по Ф-ра 00370-25",
    "dolgува": 0,
    "pobaruva": 492172.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "12.12.2025",
    "desc": "Ф-ра 00371-25",
    "dolgува": 5126424.0,
    "pobaruva": 0,
    "valuta": "26.01.2026",
    "status": "paid"
  },
  {
    "date": "12.12.2025",
    "desc": "Аванс по Ф-ра 00371-25",
    "dolgува": 0,
    "pobaruva": 828309.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.12.2025",
    "desc": "Ф-ра 00372-25",
    "dolgува": 2621124.0,
    "pobaruva": 0,
    "valuta": "31.01.2026",
    "status": "paid"
  },
  {
    "date": "17.12.2025",
    "desc": "Аванс по Ф-ра 00372-25",
    "dolgува": 0,
    "pobaruva": 442235.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "26.12.2025",
    "desc": "Уплата по Ф-ра 00367-25",
    "dolgува": 0,
    "pobaruva": 3350615.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.01.2026",
    "desc": "Уплата по Ф-ра 00368-25",
    "dolgува": 0,
    "pobaruva": 2801963.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.01.2026",
    "desc": "Ф-ра 00376-26",
    "dolgува": 4877226.0,
    "pobaruva": 0,
    "valuta": "08.03.2026",
    "status": "paid"
  },
  {
    "date": "22.01.2026",
    "desc": "Аванс по Ф-ра 00376-26",
    "dolgува": 0,
    "pobaruva": 787299.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "23.01.2026",
    "desc": "Уплата по Ф-ра 00370-25",
    "dolgува": 0,
    "pobaruva": 2461628.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "23.01.2026",
    "desc": "Уплата по Ф-ра 00371-25",
    "dolgува": 0,
    "pobaruva": 4298115.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.01.2026",
    "desc": "Уплата по Ф-ра 00372-25",
    "dolgува": 0,
    "pobaruva": 2178889.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.01.2026",
    "desc": "Ф-ра 00379-26",
    "dolgува": 2834920.0,
    "pobaruva": 0,
    "valuta": "16.03.2026",
    "status": "paid"
  },
  {
    "date": "30.01.2026",
    "desc": "Аванс по Ф-ра 00379-26",
    "dolgува": 0,
    "pobaruva": 476019.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.02.2026",
    "desc": "Ф-ра 00380-26",
    "dolgува": 3291446.0,
    "pobaruva": 0,
    "valuta": "22.03.2026",
    "status": "paid"
  },
  {
    "date": "05.02.2026",
    "desc": "Аванс по Ф-ра 00380-26",
    "dolgува": 0,
    "pobaruva": 546648.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "12.02.2026",
    "desc": "Ф-ра 00381-26",
    "dolgува": 3213637.0,
    "pobaruva": 0,
    "valuta": "29.03.2026",
    "status": "paid"
  },
  {
    "date": "12.02.2026",
    "desc": "Аванс по Ф-ра 00381-26",
    "dolgува": 0,
    "pobaruva": 534918.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "12.02.2026",
    "desc": "Ф-ра 00382-26",
    "dolgува": 4088350.0,
    "pobaruva": 0,
    "valuta": "29.03.2026",
    "status": "paid"
  },
  {
    "date": "12.02.2026",
    "desc": "Аванс по Ф-ра 00382-26",
    "dolgува": 0,
    "pobaruva": 657420.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "13.02.2026",
    "desc": "Ф-ра 00383-26",
    "dolgува": 2726797.0,
    "pobaruva": 0,
    "valuta": "30.03.2026",
    "status": "paid"
  },
  {
    "date": "13.02.2026",
    "desc": "Аванс по Ф-ра 00383-26",
    "dolgува": 0,
    "pobaruva": 459674.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "27.02.2026",
    "desc": "Ф-ра 00386-26",
    "dolgува": 3641712.0,
    "pobaruva": 0,
    "valuta": "13.04.2026",
    "status": "paid"
  },
  {
    "date": "27.02.2026",
    "desc": "Аванс по Ф-ра 00386-26",
    "dolgува": 0,
    "pobaruva": 594996.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "06.03.2026",
    "desc": "Ф-ра 00388-26",
    "dolgува": 2689968.0,
    "pobaruva": 0,
    "valuta": "20.04.2026",
    "status": "paid"
  },
  {
    "date": "06.03.2026",
    "desc": "Аванс по Ф-ра 00388-26",
    "dolgува": 0,
    "pobaruva": 454417.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "06.03.2026",
    "desc": "Уплата по Ф-ра 00376-26",
    "dolgува": 0,
    "pobaruva": 4089927.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.03.2026",
    "desc": "Уплата по Ф-ра 00379-26",
    "dolgува": 0,
    "pobaruva": 2358901.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "23.03.2026",
    "desc": "Уплата по Ф-ра 00380-26",
    "dolgува": 0,
    "pobaruva": 2744798.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.03.2026",
    "desc": "Уплата по Ф-ра 00381-26",
    "dolgува": 0,
    "pobaruva": 2678719.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.03.2026",
    "desc": "Уплата по Ф-ра 00382-26",
    "dolgува": 0,
    "pobaruva": 3430930.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.03.2026",
    "desc": "Уплата по Ф-ра 00383-26",
    "dolgува": 0,
    "pobaruva": 2267123.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "15.04.2026",
    "desc": "Уплата по Ф-ра 00386-26",
    "dolgува": 0,
    "pobaruva": 3046716.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.04.2026",
    "desc": "Ф-ра 00390-26",
    "dolgува": 2858202.0,
    "pobaruva": 0,
    "valuta": "01.06.2026",
    "status": "paid"
  },
  {
    "date": "17.04.2026",
    "desc": "Аванс по Ф-ра 00390-26",
    "dolgува": 0,
    "pobaruva": 480007.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.04.2026",
    "desc": "Уплата по Ф-ра 00388-26",
    "dolgува": 0,
    "pobaruva": 2235551.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "21.04.2026",
    "desc": "Ф-ра 00391-26",
    "dolgува": 2788733.0,
    "pobaruva": 0,
    "valuta": "05.06.2026",
    "status": "paid"
  },
  {
    "date": "21.04.2026",
    "desc": "Аванс по Ф-ра 00391-26",
    "dolgува": 0,
    "pobaruva": 469559.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "04.05.2026",
    "desc": "Ф-ра 00394-26",
    "dolgува": 4040082.0,
    "pobaruva": 0,
    "valuta": "18.06.2026",
    "status": "paid"
  },
  {
    "date": "04.05.2026",
    "desc": "Аванс по Ф-ра 00394-26",
    "dolgува": 0,
    "pobaruva": 658448.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "12.05.2026",
    "desc": "Ф-ра 00395-26",
    "dolgува": 4825034.0,
    "pobaruva": 0,
    "valuta": "26.06.2026",
    "status": "paid"
  },
  {
    "date": "12.05.2026",
    "desc": "Аванс по Ф-ра 00395-26",
    "dolgува": 0,
    "pobaruva": 788841.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "18.05.2026",
    "desc": "Ф-ра 00398-26",
    "dolgува": 2561362.0,
    "pobaruva": 0,
    "valuta": "02.07.2026",
    "status": "paid"
  },
  {
    "date": "12.05.2026",
    "desc": "Аванс по Ф-ра 00395-26",
    "dolgува": 0,
    "pobaruva": 434873.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "27.05.2026",
    "desc": "Ф-ра 00399-26",
    "dolgува": 2655913.0,
    "pobaruva": 0,
    "valuta": "11.07.2026",
    "status": "paid"
  },
  {
    "date": "27.05.2026",
    "desc": "Аванс по Ф-ра 00399-26",
    "dolgува": 0,
    "pobaruva": 450253.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "01.06.2026",
    "desc": "Уплата по Ф-ра 00390-26",
    "dolgува": 0,
    "pobaruva": 2378195.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "04.06.2026",
    "desc": "Уплата по Ф-ра 00391-26",
    "dolgува": 0,
    "pobaruva": 2319174.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "08.06.2026",
    "desc": "Ф-ра 00400-26",
    "dolgува": 4808299.0,
    "pobaruva": 0,
    "valuta": "23.07.2026",
    "status": "paid"
  },
  {
    "date": "08.06.2026",
    "desc": "Аванс по Ф-ра 00400-26",
    "dolgува": 0,
    "pobaruva": 792422.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "18.06.2026",
    "desc": "Уплата по Ф-ра 00394-26",
    "dolgува": 0,
    "pobaruva": 3381634.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "24.06.2026",
    "desc": "Ф-ра 00402-26",
    "dolgува": 2796720.0,
    "pobaruva": 0,
    "valuta": "08.08.2026",
    "status": "paid"
  },
  {
    "date": "24.06.2026",
    "desc": "Аванс по Ф-ра 00402-26",
    "dolgува": 0,
    "pobaruva": 472305.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.06.2026",
    "desc": "Ф-ра 00403-26",
    "dolgува": 3794164.0,
    "pobaruva": 0,
    "valuta": "09.08.2026",
    "status": "paid"
  },
  {
    "date": "25.06.2026",
    "desc": "Аванс по Ф-ра 00403-26",
    "dolgува": 0,
    "pobaruva": 628386.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "26.06.2026",
    "desc": "Уплата по Ф-ра 00395-26",
    "dolgува": 0,
    "pobaruva": 4036193.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "01.07.2026",
    "desc": "Уплата по Ф-ра 00398-26",
    "dolgува": 0,
    "pobaruva": 2126489.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.07.2026",
    "desc": "Уплата по Ф-ра 00399-26",
    "dolgува": 0,
    "pobaruva": 2205660.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "23.07.2026",
    "desc": "Уплата по Ф-ра 00400-26",
    "dolgува": 0,
    "pobaruva": 4015877.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.08.2026",
    "desc": "Уплата по Ф-ра 00402-26",
    "dolgува": 0,
    "pobaruva": 2324415.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.08.2026",
    "desc": "Уплата по Ф-ра 00403-26",
    "dolgува": 0,
    "pobaruva": 3165778.0,
    "valuta": "",
    "status": ""
  }
];

const CARD_IMPERIJAL = [
  {
    "date": "24.01.2025",
    "desc": "Ф-ра 00324-25",
    "dolgува": 5622662.0,
    "pobaruva": 0,
    "valuta": "10.03.2025",
    "status": "paid"
  },
  {
    "date": "24.01.2025",
    "desc": "Аванс по Ф-ра 00324-25",
    "dolgува": 0,
    "pobaruva": 912857.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "06.02.2025",
    "desc": "Ф-ра 00326-25",
    "dolgува": 2570760.0,
    "pobaruva": 0,
    "valuta": "23.03.2025",
    "status": "paid"
  },
  {
    "date": "06.02.2025",
    "desc": "Аванс по Ф-ра 00326-25",
    "dolgува": 0,
    "pobaruva": 426048.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "24.02.2025",
    "desc": "Ф-ра 00328-25",
    "dolgува": 7390771.0,
    "pobaruva": 0,
    "valuta": "10.04.2025",
    "status": "paid"
  },
  {
    "date": "24.02.2025",
    "desc": "Аванс по Ф-ра 00328-25",
    "dolgува": 0,
    "pobaruva": 1218143.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.03.2025",
    "desc": "Ф-ра 00331-25",
    "dolgува": 3330178.0,
    "pobaruva": 0,
    "valuta": "19.04.2025",
    "status": "paid"
  },
  {
    "date": "05.03.2025",
    "desc": "Аванс по Ф-ра 00331-25",
    "dolgува": 0,
    "pobaruva": 550142.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.03.2025",
    "desc": "Ф-ра 00332-25",
    "dolgува": 7815360.0,
    "pobaruva": 0,
    "valuta": "24.04.2025",
    "status": "paid"
  },
  {
    "date": "10.03.2025",
    "desc": "Аванс по Ф-ра 00332-25",
    "dolgува": 0,
    "pobaruva": 1282321.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "13.03.2025",
    "desc": "Ф-ра 00333-25",
    "dolgува": 2438706.0,
    "pobaruva": 0,
    "valuta": "27.04.2025",
    "status": "paid"
  },
  {
    "date": "13.03.2025",
    "desc": "Аванс по Ф-ра 00333-25",
    "dolgува": 0,
    "pobaruva": 407479.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.03.2025",
    "desc": "Уплата по Ф-ра 00324-25",
    "dolgува": 0,
    "pobaruva": 1709805.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.03.2025",
    "desc": "Уплата по Ф-ра 00324-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.03.2025",
    "desc": "Уплата по Ф-ра 00324-25",
    "dolgува": 0,
    "pobaruva": 2000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "21.03.2025",
    "desc": "Ф-ра 00336-25",
    "dolgува": 2319697.0,
    "pobaruva": 0,
    "valuta": "05.05.2025",
    "status": "paid"
  },
  {
    "date": "21.03.2025",
    "desc": "Аванс по Ф-ра 00336-25",
    "dolgува": 0,
    "pobaruva": 388663.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "03.04.2025",
    "desc": "Уплата по Ф-ра 00326-25",
    "dolgува": 0,
    "pobaruva": 1144712.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "04.04.2025",
    "desc": "Уплата по Ф-ра 00326-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "08.04.2025",
    "desc": "Ф-ра 00340-25",
    "dolgува": 2900441.0,
    "pobaruva": 0,
    "valuta": "23.05.2025",
    "status": "paid"
  },
  {
    "date": "08.04.2025",
    "desc": "Аванс по Ф-ра 00340-25",
    "dolgува": 0,
    "pobaruva": 482574.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.04.2025",
    "desc": "Уплата по Ф-ра 00328-25",
    "dolgува": 0,
    "pobaruva": 2000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.04.2025",
    "desc": "Уплата по Ф-ра 00328-25",
    "dolgува": 0,
    "pobaruva": 1172628.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "24.04.2025",
    "desc": "Уплата по Ф-ра 00328-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.04.2025",
    "desc": "Уплата по Ф-ра 00328-25",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.04.2025",
    "desc": "Уплата по Ф-ра 00328-25",
    "dolgува": 0,
    "pobaruva": 1500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.04.2025",
    "desc": "Ф-ра 00341-25",
    "dolgува": 7089110.0,
    "pobaruva": 0,
    "valuta": "09.06.2025",
    "status": "paid"
  },
  {
    "date": "25.04.2025",
    "desc": "Аванс по Ф-ра 00341-25",
    "dolgува": 0,
    "pobaruva": 1144208.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "06.05.2025",
    "desc": "Уплата по Ф-ра 00331-25",
    "dolgува": 0,
    "pobaruva": 1780036.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.05.2025",
    "desc": "Уплата по Ф-ра 00331-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.05.2025",
    "desc": "Уплата по Ф-ра 00332-25",
    "dolgува": 0,
    "pobaruva": 1533039.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "13.05.2025",
    "desc": "Ф-ра 00342-25",
    "dolgува": 2637000.0,
    "pobaruva": 0,
    "valuta": "27.06.2025",
    "status": "paid"
  },
  {
    "date": "13.05.2025",
    "desc": "Аванс по Ф-ра 00342-25",
    "dolgува": 0,
    "pobaruva": 435623.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "19.05.2025",
    "desc": "Ф-ра 00344-25",
    "dolgува": 5428168.0,
    "pobaruva": 0,
    "valuta": "03.07.2025",
    "status": "paid"
  },
  {
    "date": "19.05.2025",
    "desc": "Аванс по Ф-ра 00344-25",
    "dolgува": 0,
    "pobaruva": 600000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "19.05.2025",
    "desc": "Аванс по Ф-ра 00344-25",
    "dolgува": 0,
    "pobaruva": 284233.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.05.2025",
    "desc": "Уплата по Ф-ра 00332-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.05.2025",
    "desc": "Уплата по Ф-ра 00332-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "23.05.2025",
    "desc": "Уплата по Ф-ра 00332-25",
    "dolgува": 0,
    "pobaruva": 2600000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "23.05.2025",
    "desc": "Уплата по Ф-ра 00332-25",
    "dolgува": 0,
    "pobaruva": 400000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "04.06.2025",
    "desc": "Уплата по Ф-ра 00333-25",
    "dolgува": 0,
    "pobaruva": 1031227.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.06.2025",
    "desc": "Уплата по Ф-ра 00333-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "09.06.2025",
    "desc": "Уплата по Ф-ра 00336-25",
    "dolgува": 0,
    "pobaruva": 1931034.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.06.2025",
    "desc": "Ф-ра 00346-25",
    "dolgува": 2976300.0,
    "pobaruva": 0,
    "valuta": "01.08.2025",
    "status": "paid"
  },
  {
    "date": "17.06.2025",
    "desc": "Аванс по Ф-ра 00346-25",
    "dolgува": 0,
    "pobaruva": 489252.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.06.2025",
    "desc": "Ф-ра 00348-25",
    "dolgува": 6643386.0,
    "pobaruva": 0,
    "valuta": "04.08.2025",
    "status": "paid"
  },
  {
    "date": "20.06.2025",
    "desc": "Аванс по Ф-ра 00348-25",
    "dolgува": 0,
    "pobaruva": 1102839.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "23.06.2025",
    "desc": "Уплата по Ф-ра 00340-25",
    "dolgува": 0,
    "pobaruva": 1517867.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "24.06.2025",
    "desc": "Уплата по Ф-ра 00340-25",
    "dolgува": 0,
    "pobaruva": 900000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "26.06.2025",
    "desc": "Уплата по Ф-ра 00341-25",
    "dolgува": 0,
    "pobaruva": 2004902.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "03.07.2025",
    "desc": "Уплата по Ф-ра 00341-25",
    "dolgува": 0,
    "pobaruva": 940000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "04.07.2025",
    "desc": "Уплата по Ф-ра 00341-25",
    "dolgува": 0,
    "pobaruva": 1700000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.07.2025",
    "desc": "Уплата по Ф-ра 00341-25",
    "dolgува": 0,
    "pobaruva": 1300000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.07.2025",
    "desc": "Уплата по Ф-ра 00342-25",
    "dolgува": 0,
    "pobaruva": 1201377.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "15.07.2025",
    "desc": "Уплата по Ф-ра 00342-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "16.07.2025",
    "desc": "Уплата по Ф-ра 00344-25",
    "dolgува": 0,
    "pobaruva": 124587.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "16.07.2025",
    "desc": "Ф-ра 00350-25",
    "dolgува": 3156014.0,
    "pobaruva": 0,
    "valuta": "30.08.2025",
    "status": "paid"
  },
  {
    "date": "16.07.2025",
    "desc": "Аванс по Ф-ра 00350-25",
    "dolgува": 0,
    "pobaruva": 513486.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "18.07.2025",
    "desc": "Уплата по Ф-ра 00344-25",
    "dolgува": 0,
    "pobaruva": 1419348.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "24.07.2025",
    "desc": "Уплата по Ф-ра 00344-25",
    "dolgува": 0,
    "pobaruva": 1500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.07.2025",
    "desc": "Уплата по Ф-ра 00344-25",
    "dolgува": 0,
    "pobaruva": 1500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "06.08.2025",
    "desc": "Ф-ра 00351-25",
    "dolgува": 8231226.0,
    "pobaruva": 0,
    "valuta": "20.09.2025",
    "status": "paid"
  },
  {
    "date": "06.08.2025",
    "desc": "Аванс по Ф-ра 00351-25",
    "dolgува": 0,
    "pobaruva": 1375069.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.08.2025",
    "desc": "Уплата по Ф-ра 00346-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.08.2025",
    "desc": "Уплата по Ф-ра 00346-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.08.2025",
    "desc": "Уплата по Ф-ра 00346-25",
    "dolgува": 0,
    "pobaruva": 487048.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.08.2025",
    "desc": "Уплата по Ф-ра 00348-25",
    "dolgува": 0,
    "pobaruva": 1540547.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "01.09.2025",
    "desc": "Ф-ра 00354-25",
    "dolgува": 3329544.0,
    "pobaruva": 0,
    "valuta": "16.10.2025",
    "status": "paid"
  },
  {
    "date": "01.09.2025",
    "desc": "Аванс по Ф-ра 00354-25",
    "dolgува": 0,
    "pobaruva": 537041.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "25.08.2025",
    "desc": "Уплата по Ф-ра 00348-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "26.08.2025",
    "desc": "Уплата по Ф-ра 00348-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "29.08.2025",
    "desc": "Уплата по Ф-ра 00348-25",
    "dolgува": 0,
    "pobaruva": 1200000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "29.08.2025",
    "desc": "Уплата по Ф-ра 00348-25",
    "dolgува": 0,
    "pobaruva": 800000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "12.09.2025",
    "desc": "Ф-ра 00356-25",
    "dolgува": 2816047.0,
    "pobaruva": 0,
    "valuta": "27.10.2025",
    "status": "paid"
  },
  {
    "date": "12.09.2025",
    "desc": "Аванс по Ф-ра 00356-25",
    "dolgува": 0,
    "pobaruva": 468135.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "16.09.2025",
    "desc": "Ф-ра 00357-25",
    "dolgува": 7503406.0,
    "pobaruva": 0,
    "valuta": "31.10.2025",
    "status": "paid"
  },
  {
    "date": "16.09.2025",
    "desc": "Аванс по Ф-ра 00357-25",
    "dolgува": 0,
    "pobaruva": 1249771.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "16.09.2025",
    "desc": "Уплата по Ф-ра 00350-25",
    "dolgува": 0,
    "pobaruva": 642528.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "18.09.2025",
    "desc": "Уплата по Ф-ра 00350-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "19.09.2025",
    "desc": "Уплата по Ф-ра 00350-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "26.09.2025",
    "desc": "Ф-ра 00360-25",
    "dolgува": 4184556.0,
    "pobaruva": 0,
    "valuta": "10.11.2025",
    "status": "paid"
  },
  {
    "date": "26.09.2025",
    "desc": "Аванс по Ф-ра 00360-25",
    "dolgува": 0,
    "pobaruva": 681445.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "02.10.2025",
    "desc": "Уплата по Ф-ра 00351-25",
    "dolgува": 0,
    "pobaruva": 856157.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "03.10.2025",
    "desc": "Уплата по Ф-ра 00351-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.10.2025",
    "desc": "Уплата по Ф-ра 00351-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.10.2025",
    "desc": "Ф-ра 00363-25",
    "dolgува": 2799632.0,
    "pobaruva": 0,
    "valuta": "21.11.2025",
    "status": "paid"
  },
  {
    "date": "07.10.2025",
    "desc": "Аванс по Ф-ра 00363-25",
    "dolgува": 0,
    "pobaruva": 454175.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "08.10.2025",
    "desc": "Уплата по Ф-ра 00351-25",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "09.10.2025",
    "desc": "Уплата по Ф-ра 00351-25",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.10.2025",
    "desc": "Уплата по Ф-ра 00351-25",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "15.10.2025",
    "desc": "Уплата по Ф-ра 00351-25",
    "dolgува": 0,
    "pobaruva": 1500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.10.2025",
    "desc": "Уплата по Ф-ра 00351-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.10.2025",
    "desc": "Ф-ра 00366-25",
    "dolgува": 2696434.0,
    "pobaruva": 0,
    "valuta": "14.12.2025",
    "status": "paid"
  },
  {
    "date": "30.10.2025",
    "desc": "Аванс по Ф-ра 00366-25",
    "dolgува": 0,
    "pobaruva": 454143.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "19.11.2025",
    "desc": "Уплата по Ф-ра 00354-25",
    "dolgува": 0,
    "pobaruva": 2792503.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "04.12.2025",
    "desc": "Ф-ра 00369-25",
    "dolgува": 4771904.0,
    "pobaruva": 0,
    "valuta": "18.01.2026",
    "status": "paid"
  },
  {
    "date": "04.12.2025",
    "desc": "Аванс по Ф-ра 00369-25",
    "dolgува": 0,
    "pobaruva": 776051.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.12.2025",
    "desc": "Уплата по Ф-ра 00356-25",
    "dolgува": 0,
    "pobaruva": 2347912.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.12.2025",
    "desc": "Уплата по Ф-ра 00360-25",
    "dolgува": 0,
    "pobaruva": 3503111.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.12.2025",
    "desc": "Уплата по Ф-ра 00357-25",
    "dolgува": 0,
    "pobaruva": 148977.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "09.12.2025",
    "desc": "Уплата по Ф-ра 00363-25",
    "dolgува": 0,
    "pobaruva": 2345457.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "09.12.2025",
    "desc": "Уплата по Ф-ра 00357-25",
    "dolgува": 0,
    "pobaruva": 256090.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "16.12.2025",
    "desc": "Уплата по Ф-ра 00357-25",
    "dolgува": 0,
    "pobaruva": 1503111.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.12.2025",
    "desc": "Уплата по Ф-ра 00357-25",
    "dolgува": 0,
    "pobaruva": 2000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "19.12.2025",
    "desc": "Ф-ра 00373-25",
    "dolgува": 6460149.0,
    "pobaruva": 0,
    "valuta": "02.02.2026",
    "status": "paid"
  },
  {
    "date": "19.12.2025",
    "desc": "Аванс по Ф-ра 00373-25",
    "dolgува": 0,
    "pobaruva": 1061698.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "19.12.2025",
    "desc": "Ф-ра 00374-25",
    "dolgува": 2363609.0,
    "pobaruva": 0,
    "valuta": "02.02.2026",
    "status": "paid"
  },
  {
    "date": "19.12.2025",
    "desc": "Аванс по Ф-ра 00374-25",
    "dolgува": 0,
    "pobaruva": 400783.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "31.12.2025",
    "desc": "Уплата по Ф-ра 00357-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "02.01.2026",
    "desc": "Уплата по Ф-ра 00357-25",
    "dolgува": 0,
    "pobaruva": 1345457.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "23.01.2026",
    "desc": "Ф-ра 00377-26",
    "dolgува": 6020439.0,
    "pobaruva": 0,
    "valuta": "09.03.2026",
    "status": "paid"
  },
  {
    "date": "23.01.2026",
    "desc": "Аванс по Ф-ра 00377-26",
    "dolgува": 0,
    "pobaruva": 981538.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "23.01.2026",
    "desc": "Ф-ра 00378-26",
    "dolgува": 3129912.0,
    "pobaruva": 0,
    "valuta": "09.03.2026",
    "status": "paid"
  },
  {
    "date": "23.01.2026",
    "desc": "Аванс по Ф-ра 00378-26",
    "dolgува": 0,
    "pobaruva": 521125.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.01.2026",
    "desc": "Уплата по Ф-ра 00366-25",
    "dolgува": 0,
    "pobaruva": 1415000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.01.2026",
    "desc": "Уплата по Ф-ра 00366-25",
    "dolgува": 0,
    "pobaruva": 827291.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "09.02.2026",
    "desc": "Уплата по Ф-ра 00369-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.02.2026",
    "desc": "Уплата по Ф-ра 00369-25",
    "dolgува": 0,
    "pobaruva": 1400000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "12.02.2026",
    "desc": "Уплата по Ф-ра 00369-25",
    "dolgува": 0,
    "pobaruva": 1595853.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.02.2026",
    "desc": "Ф-ра 00384-26",
    "dolgува": 2842230.0,
    "pobaruva": 0,
    "valuta": "06.04.2026",
    "status": "paid"
  },
  {
    "date": "20.02.2026",
    "desc": "Аванс по Ф-ра 00384-26",
    "dolgува": 0,
    "pobaruva": 476552.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.02.2026",
    "desc": "Уплата по Ф-ра 00374-25",
    "dolgува": 0,
    "pobaruva": 1962826.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "26.02.2026",
    "desc": "Уплата по Ф-ра 00373-25",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "02.03.2026",
    "desc": "Ф-ра 00387-26",
    "dolgува": 4676081.0,
    "pobaruva": 0,
    "valuta": "16.04.2026",
    "status": "paid"
  },
  {
    "date": "03.03.2026",
    "desc": "Аванс по Ф-ра 00387-26",
    "dolgува": 0,
    "pobaruva": 775322.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "04.03.2026",
    "desc": "Уплата по Ф-ра 00373-25",
    "dolgува": 0,
    "pobaruva": 1298451.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.03.2026",
    "desc": "Уплата по Ф-ра 00373-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "06.03.2026",
    "desc": "Уплата по Ф-ра 00373-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "09.03.2026",
    "desc": "Уплата по Ф-ра 00373-25",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.03.2026",
    "desc": "Уплата по Ф-ра 00373-25",
    "dolgува": 0,
    "pobaruva": 600000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "19.03.2026",
    "desc": "Ф-ра 00389-26",
    "dolgува": 3873192.0,
    "pobaruva": 0,
    "valuta": "03.05.2026",
    "status": "paid"
  },
  {
    "date": "19.03.2026",
    "desc": "Аванс по Ф-ра 00389-26",
    "dolgува": 0,
    "pobaruva": 644582.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "15.04.2026",
    "desc": "Уплата по Ф-ра 00378-26",
    "dolgува": 0,
    "pobaruva": 1208787.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "15.04.2026",
    "desc": "Уплата по Ф-ра 00378-26",
    "dolgува": 0,
    "pobaruva": 1100000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "16.04.2026",
    "desc": "Уплата по Ф-ра 00378-26",
    "dolgува": 0,
    "pobaruva": 300000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.04.2026",
    "desc": "Уплата по Ф-ра 00377-26",
    "dolgува": 0,
    "pobaruva": 1038901.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "20.04.2026",
    "desc": "Уплата по Ф-ра 00377-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "21.04.2026",
    "desc": "Уплата по Ф-ра 00377-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.04.2026",
    "desc": "Уплата по Ф-ра 00377-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "24.04.2026",
    "desc": "Уплата по Ф-ра 00377-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "27.04.2026",
    "desc": "Ф-ра 00392-26",
    "dolgува": 3739917.0,
    "pobaruva": 0,
    "valuta": "11.06.2026",
    "status": "paid"
  },
  {
    "date": "27.04.2026",
    "desc": "Аванс по Ф-ра 00392-26",
    "dolgува": 0,
    "pobaruva": 624390.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "05.05.2026",
    "desc": "Уплата по Ф-ра 00384-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.05.2026",
    "desc": "Уплата по Ф-ра 00384-26",
    "dolgува": 0,
    "pobaruva": 365678.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "08.05.2026",
    "desc": "Уплата по Ф-ра 00384-26",
    "dolgува": 0,
    "pobaruva": 300000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "08.05.2026",
    "desc": "Уплата по Ф-ра 00384-26",
    "dolgува": 0,
    "pobaruva": 700000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "13.05.2026",
    "desc": "Ф-ра 00396-26",
    "dolgува": 3063680.0,
    "pobaruva": 0,
    "valuta": "27.06.2026",
    "status": "paid"
  },
  {
    "date": "13.05.2026",
    "desc": "Аванс по Ф-ра 00396-26",
    "dolgува": 0,
    "pobaruva": 512778.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "18.05.2026",
    "desc": "Ф-ра 00397-26",
    "dolgува": 6600718.0,
    "pobaruva": 0,
    "valuta": "02.07.2026",
    "status": "paid"
  },
  {
    "date": "18.05.2026",
    "desc": "Аванс по Ф-ра 00397-26",
    "dolgува": 0,
    "pobaruva": 1084865.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "08.06.2026",
    "desc": "Уплата по Ф-ра 00387-26",
    "dolgува": 0,
    "pobaruva": 900759.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "09.06.2026",
    "desc": "Уплата по Ф-ра 00387-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.06.2026",
    "desc": "Уплата по Ф-ра 00387-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "11.06.2026",
    "desc": "Уплата по Ф-ра 00387-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "15.06.2026",
    "desc": "Ф-ра 00401-26",
    "dolgува": 3437880.0,
    "pobaruva": 0,
    "valuta": "30.07.2026",
    "status": "unpaid"
  },
  {
    "date": "12.06.2026",
    "desc": "Аванс по Ф-ра 00401-26",
    "dolgува": 0,
    "pobaruva": 572585.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.06.2026",
    "desc": "Уплата по Ф-ра 00389-26",
    "dolgува": 0,
    "pobaruva": 228610.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.06.2026",
    "desc": "Уплата по Ф-ра 00389-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.06.2026",
    "desc": "Уплата по Ф-ра 00389-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "24.06.2026",
    "desc": "Уплата по Ф-ра 00389-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "24.06.2026",
    "desc": "Уплата по Ф-ра 00389-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "26.06.2026",
    "desc": "Ф-ра 00404-26",
    "dolgува": 7847449.0,
    "pobaruva": 0,
    "valuta": "10.08.2026",
    "status": "unpaid"
  },
  {
    "date": "26.06.2026",
    "desc": "Аванс по Ф-ра 00404-26",
    "dolgува": 0,
    "pobaruva": 1291564.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "02.07.2026",
    "desc": "Ф-ра 00405-26",
    "dolgува": 6300067.0,
    "pobaruva": 0,
    "valuta": "16.08.2026",
    "status": "unpaid"
  },
  {
    "date": "02.07.2026",
    "desc": "Аванс по Ф-ра 00405-26",
    "dolgува": 0,
    "pobaruva": 1032084.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "02.07.2026",
    "desc": "Уплата по Ф-ра 00389-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "03.07.2026",
    "desc": "Уплата по Ф-ра 00389-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.07.2026",
    "desc": "Уплата по Ф-ра 00392-26",
    "dolgува": 0,
    "pobaruva": 300000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "08.07.2026",
    "desc": "Уплата по Ф-ра 00392-26",
    "dolgува": 0,
    "pobaruva": 415527.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.07.2026",
    "desc": "Уплата по Ф-ра 00392-26",
    "dolgува": 0,
    "pobaruva": 400000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "13.07.2026",
    "desc": "Уплата по Ф-ра 00392-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "13.07.2026",
    "desc": "Уплата по Ф-ра 00392-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.07.2026",
    "desc": "Уплата по Ф-ра 00392-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "15.07.2026",
    "desc": "Уплата по Ф-ра 00392-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "22.07.2026",
    "desc": "Уплата по Ф-ра 00396-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "27.07.2026",
    "desc": "Уплата по Ф-ра 00396-26",
    "dolgува": 0,
    "pobaruva": 550902.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "28.07.2026",
    "desc": "Уплата по Ф-ра 00396-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "30.07.2026",
    "desc": "Ф-ра 00407-26",
    "dolgува": 2649465.0,
    "pobaruva": 0,
    "valuta": "",
    "status": "unpaid"
  },
  {
    "date": "30.07.2026",
    "desc": "Аванс по Ф-ра 00407-26",
    "dolgува": 0,
    "pobaruva": 446614.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "06.08.2026",
    "desc": "Уплата по Ф-ра 00397-26",
    "dolgува": 0,
    "pobaruva": 1015853.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "07.08.2026",
    "desc": "Уплата по Ф-ра 00397-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "10.08.2026",
    "desc": "Уплата по Ф-ра 00397-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "13.08.2026",
    "desc": "Уплата по Ф-ра 00397-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "14.08.2026",
    "desc": "Уплата по Ф-ра 00397-26",
    "dolgува": 0,
    "pobaruva": 1000000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.08.2026",
    "desc": "Уплата по Ф-ра 00397-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  },
  {
    "date": "17.08.2026",
    "desc": "Уплата по Ф-ра 00397-26",
    "dolgува": 0,
    "pobaruva": 500000.0,
    "valuta": "",
    "status": ""
  }
];

const USERS = {
  admin:   { password: "senfood2026", role: "Admin",   name: "Ersin" },
  manager: { password: "manager2026", role: "Manager", name: "Andrijana" },
};
