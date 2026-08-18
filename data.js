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
