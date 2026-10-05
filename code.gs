/************************************************************
 * کائنات‌چی — VIP Backend
 * Google Apps Script
 *
 * نسخه کامل
 ************************************************************/
/* =========================================================
   CONFIG
========================================================= */
const VIP_SHEET_ID =
  '1TpljGwyRpHcxyyR6Zm1mbT2CPJBvFeHCvD3QH34BuGs';
const CUSTOMERS_SHEET =
  'مشتریان';
const HISTORY_SHEET =
  'سوابق نوبت ها';
const PAYMENTS_SHEET =
  'پرداخت ها';
const TOKENS_SHEET =
  'توکن ها';
const CONNECTION_CODES_SHEET =
  'کدهای اتصال VIP';
const TOKEN_EXPIRY_HOURS =
  10;
const ALLOWED_DISCOUNTS =
  [3, 5, 7];
const VIP_ACTIVE_STATUS =
  'فعال';
const CONNECTION_CODE_ACTIVE_STATUS =
  'صادرشده';
const CONNECTION_CODE_USED_STATUS =
  'استفاده‌شده';
const CONNECTION_CODE_REVOKED_STATUS =
  'باطل‌شده';
/* =========================================================
   WEB APP
========================================================= */
function doGet(e) {
  const page =
    e &&
    e.parameter &&
    e.parameter.page
      ? e.parameter.page
      : '';
  if (page === 'admin') {
    return HtmlService
      .createHtmlOutputFromFile('Admin')
      .setTitle(
        'پنل مدیریت VIP کائنات‌چی'
      )
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );
  }
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle(
      'پنل VIP کائنات‌چی'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}
/* =========================================================
   SHEET HELPERS
========================================================= */
function getSheetOrThrow_(sheetName) {
  const ss =
    SpreadsheetApp.openById(
      VIP_SHEET_ID
    );
  const sheet =
    ss.getSheetByName(
      sheetName
    );
  if (!sheet) {
    throw new Error(
      'شیت پیدا نشد: ' +
      sheetName
    );
  }
  return sheet;
}
/* =========================================================
   GENERAL HELPERS
========================================================= */
function isVIPActive_(status) {
  return (
    String(status || '')
      .trim() ===
    VIP_ACTIVE_STATUS
  );
}
function normalizeString_(value) {
  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }
  return String(value).trim();
}
function generateRandomCode_(
  prefix,
  length
) {
  const chars =
    'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let result =
    prefix || '';
  for (
    let i = 0;
    i < length;
    i++
  ) {
    result +=
      chars.charAt(
        Math.floor(
          Math.random() *
          chars.length
        )
      );
  }
  return result;
}
/* =========================================================
   CUSTOMER ID RESOLVER
========================================================= */
/*
 * این تابع ورودی‌های مختلفی را که ممکن است
 * Admin.html ارسال کند به ID واقعی مشتری تبدیل می‌کند.
 *
 * قابل قبول:
 *
 * VIP001
 *
 * {
 *   id: 'VIP001'
 * }
 *
 * {
 *   customerId: 'VIP001'
 * }
 *
 * {
 *   value: 'VIP001'
 * }
 *
 * همچنین اگر آبجکت نام مشتری داشته باشد،
 * از روی نام مشتری ID پیدا می‌شود.
 */
function extractCustomerId_(value) {
  /*
   * اگر مقدار مستقیماً یک شناسه متنی/عددی باشد
   */
  if (
    value !== null &&
    value !== undefined &&
    typeof value !== 'object'
  ) {
    return normalizeString_(value);
  }
  /*
   * اگر مقدار یک آبجکت مشتری باشد،
   * ابتدا customerId را بررسی می‌کنیم.
   *
   * مهم:
   * بعضی آبجکت‌های پنل ممکن است همزمان
   * id و customerId داشته باشند.
   * id ممکن است شماره ردیف یا شناسه داخلی باشد،
   * بنابراین customerId باید اولویت داشته باشد.
   */
  if (
    value &&
    typeof value === 'object'
  ) {
    const possibleCustomerIdFields = [
      'customerId',
      'customerID',
      'vipCustomerId',
      'vipId',
      'customer_id'
    ];
    for (
      let i = 0;
      i < possibleCustomerIdFields.length;
      i++
    ) {
      const field =
        possibleCustomerIdFields[i];
      if (
        value[field] !== undefined &&
        value[field] !== null &&
        normalizeString_(value[field])
      ) {
        return normalizeString_(
          value[field]
        );
      }
    }
    /*
     * اگر customerId وجود نداشت،
     * بعد سراغ id می‌رویم.
     */
    if (
      value.id !== undefined &&
      value.id !== null
    ) {
      const idValue =
        normalizeString_(
          value.id
        );
      if (idValue) {
        return idValue;
      }
    }
    /*
     * پشتیبانی از آبجکت‌هایی که
     * مقدارشان داخل value قرار دارد.
     */
    if (
      value.value !== undefined &&
      value.value !== null
    ) {
      const valueId =
        normalizeString_(
          value.value
        );
      if (valueId) {
        return valueId;
      }
    }
    /*
     * اگر شناسه مستقیماً پیدا نشد،
     * از نام و نام خانوادگی مشتری
     * برای پیدا کردن ID واقعی استفاده می‌کنیم.
     */
    const firstName =
      normalizeString_(
        value.firstName
      );
    const lastName =
      normalizeString_(
        value.lastName
      );
    if (
      firstName ||
      lastName
    ) {
      const sheet =
        getSheetOrThrow_(
          CUSTOMERS_SHEET
        );
      const lastRow =
        sheet.getLastRow();
      if (lastRow >= 2) {
        const rows =
          sheet
            .getRange(
              2,
              1,
              lastRow - 1,
              8
            )
            .getValues();
        for (
          let i = 0;
          i < rows.length;
          i++
        ) {
          const rowFirstName =
            normalizeString_(
              rows[i][1]
            );
          const rowLastName =
            normalizeString_(
              rows[i][2]
            );
          if (
            firstName === rowFirstName &&
            lastName === rowLastName
          ) {
            return normalizeString_(
              rows[i][0]
            );
          }
        }
      }
    }
  }
  return '';
}
/* =========================================================
   FIND CUSTOMER ROW
========================================================= */
function findCustomerRow_(
  customerId
) {
  const sheet =
    getSheetOrThrow_(
      CUSTOMERS_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }
  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        8
      )
      .getValues();
  /*
   * ابتدا ID را استخراج می‌کنیم.
   */
  const target =
    extractCustomerId_(
      customerId
    );
  /*
   * جستجوی مستقیم ID
   */
  if (target) {
    for (
      let i = 0;
      i < rows.length;
      i++
    ) {
      if (
        normalizeString_(
          rows[i][0]
        ) === target
      ) {
        return i + 2;
      }
    }
  }
  /*
   * اگر ID پیدا نشد، مقدار خام را
   * به عنوان نام مشتری بررسی می‌کنیم.
   */
  const rawValue =
    normalizeString_(
      customerId
    );
  if (!rawValue) {
    return -1;
  }
  const normalizedInput =
    rawValue
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  for (
    let i = 0;
    i < rows.length;
    i++
  ) {
    const rowId =
      normalizeString_(
        rows[i][0]
      );
    const firstName =
      normalizeString_(
        rows[i][1]
      );
    const lastName =
      normalizeString_(
        rows[i][2]
      );
    const fullName =
      (
        firstName +
        ' ' +
        lastName
      )
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    if (
      normalizedInput ===
      rowId.toLowerCase()
    ) {
      return i + 2;
    }
    if (
      normalizedInput ===
      firstName.toLowerCase()
    ) {
      return i + 2;
    }
    if (
      normalizedInput ===
      lastName.toLowerCase()
    ) {
      return i + 2;
    }
    if (
      normalizedInput ===
      fullName
    ) {
      return i + 2;
    }
  }
  return -1;
}
/* =========================================================
   FIND CUSTOMER BY TELEGRAM ID
========================================================= */
function findCustomerByTelegramId_(
  telegramId
) {
  const sheet =
    getSheetOrThrow_(
      CUSTOMERS_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        8
      )
      .getValues();
  const target =
    normalizeString_(
      telegramId
    );
  for (
    let i = 0;
    i < values.length;
    i++
  ) {
    const row =
      values[i];
    if (
      normalizeString_(
        row[4]
      ) === target
    ) {
      return {
        rowNumber: i + 2,
        id:
          normalizeString_(row[0]),
        firstName:
          normalizeString_(row[1]),
        lastName:
          normalizeString_(row[2]),
        mobile:
          normalizeString_(row[3]),
        telegramId:
          normalizeString_(row[4]),
        joinedAt:
          row[5],
        bookingsCount:
          row[6],
        vipStatus:
          normalizeString_(row[7])
      };
    }
  }
  return null;
}
/* =========================================================
   CUSTOMER OBJECT
========================================================= */
function getCustomerObject_(
  row
) {
  return {
    id:
      normalizeString_(row[0]),
    firstName:
      normalizeString_(row[1]),
    lastName:
      normalizeString_(row[2]),
    mobile:
      normalizeString_(row[3]),
    telegramId:
      normalizeString_(row[4]),
    joinedAt:
      formatPersianDateTime(
        row[5]
      ),
    bookingsCount:
      row[6] || 0,
    vipStatus:
      normalizeString_(row[7])
  };
}
/* =========================================================
   ADMIN INITIAL DATA
========================================================= */
function getAdminInitialData() {
  const result = {
    success: true,
    customers: [],
    tokens: [],
    connectionCodes: [],
    settings: {},
    customersSuccess: false,
    tokensSuccess: false,
    connectionCodesSuccess: false,
    settingsSuccess: false
  };
  try {
    const customersResult =
      getVIPCustomers();
    if (
      customersResult &&
      customersResult.success
    ) {
      result.customers =
        customersResult.customers || [];
      result.customersSuccess =
        true;
    }
  } catch (err) {
    result.customersSuccess =
      false;
  }
  try {
    const tokensResult =
      getVIPTokens();
    if (
      tokensResult &&
      tokensResult.success
    ) {
      result.tokens =
        tokensResult.tokens || [];
      result.tokensSuccess =
        true;
    }
  } catch (err) {
    result.tokensSuccess =
      false;
  }
  try {
    const codesResult =
      getVIPConnectionCodes();
    if (
      codesResult &&
      codesResult.success
    ) {
      result.connectionCodes =
        codesResult.codes || [];
      result.connectionCodesSuccess =
        true;
    }
  } catch (err) {
    result.connectionCodesSuccess =
      false;
  }
  try {
    const settingsResult =
      getVIPSettings();
    if (
      settingsResult &&
      settingsResult.success
    ) {
      result.settings =
        settingsResult.settings || {};
      result.settingsSuccess =
        true;
    }
  } catch (err) {
    result.settings = {
      bookingUrl:
        'https://kaenatchi.github.io/booking/'
    };
    result.settingsSuccess =
      true;
  }
  return result;
}
function testAdminInitialData() {
  const result =
    getAdminInitialData();
  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
  return result;
}
/* =========================================================
   TELEGRAM WEB APP VALIDATION
========================================================= */
function getTelegramBotToken_() {
  const props =
    PropertiesService
      .getScriptProperties();
  const token =
    props.getProperty(
      'TELEGRAM_BOT_TOKEN'
    );
  if (!token) {
    throw new Error(
      'TELEGRAM_BOT_TOKEN در Script Properties تنظیم نشده است.'
    );
  }
  return token;
}
function bytesToHex_(
  bytes
) {
  return bytes
    .map(
      function(byte) {
        const value =
          byte < 0
            ? byte + 256
            : byte;
        return (
          '0' +
          value.toString(16)
        ).slice(-2);
      }
    )
    .join('');
}
function hexToBytes_(
  hex
) {
  const bytes = [];
  for (
    let i = 0;
    i < hex.length;
    i += 2
  ) {
    bytes.push(
      parseInt(
        hex.substr(i, 2),
        16
      )
    );
  }
  return bytes;
}
function hmacSha256Hex_(
  keyBytes,
  text
) {
  const signature =
    Utilities
      .computeHmacSha256Signature(
        text,
        keyBytes
      );
  return bytesToHex_(
    signature
  );
}
function validateTelegramInitData_(
  initData
) {
  if (!initData) {
    throw new Error(
      'Telegram initData ارسال نشده است.'
    );
  }
  const botToken =
    getTelegramBotToken_();
  const params = {};
  initData
    .split('&')
    .forEach(
      function(part) {
        const separator =
          part.indexOf('=');
        if (
          separator === -1
        ) {
          return;
        }
        const key =
          decodeURIComponent(
            part.substring(
              0,
              separator
            )
          );
        const value =
          decodeURIComponent(
            part.substring(
              separator + 1
            )
          );
        params[key] =
          value;
      }
    );
  const receivedHash =
    params.hash;
  if (!receivedHash) {
    throw new Error(
      'Telegram hash وجود ندارد.'
    );
  }
  const dataCheckString =
    Object
      .keys(params)
      .filter(
        function(key) {
          return key !== 'hash';
        }
      )
      .sort()
      .map(
        function(key) {
          return (
            key +
            '=' +
            params[key]
          );
        }
      )
      .join('\n');
  const secretKey =
    Utilities
      .computeHmacSha256Signature(
        botToken,
        'WebAppData'
      );
  const calculatedHash =
    hmacSha256Hex_(
      secretKey,
      dataCheckString
    );
  if (
    calculatedHash.toLowerCase() !==
    receivedHash.toLowerCase()
  ) {
    throw new Error(
      'Telegram initData معتبر نیست.'
    );
  }
  let user = null;
  if (params.user) {
    try {
      user =
        JSON.parse(
          params.user
        );
    } catch (err) {
      throw new Error(
        'اطلاعات کاربر Telegram قابل خواندن نیست.'
      );
    }
  }
  return {
    params:
      params,
    user:
      user
  };
}
/* =========================================================
   TELEGRAM VIP API
========================================================= */
function doPost(e) {
  try {
    const body =
      e &&
      e.postData &&
      e.postData.contents
        ? JSON.parse(
            e.postData.contents
          )
        : {};
    const initData =
      body.initData || '';
    const telegramData =
      validateTelegramInitData_(
        initData
      );
    const user =
      telegramData.user;
    if (
      !user ||
      !user.id
    ) {
      throw new Error(
        'کاربر Telegram پیدا نشد.'
      );
    }
    const customer =
      findCustomerByTelegramId_(
        String(user.id)
      );
    if (!customer) {
      return ContentService
        .createTextOutput(
          JSON.stringify({
            success: false,
            code:
              'NOT_CONNECTED',
            message:
              'این حساب Telegram هنوز به باشگاه VIP متصل نشده است.'
          })
        )
        .setMimeType(
          ContentService.MimeType.JSON
        );
    }
    if (
      !isVIPActive_(
        customer.vipStatus
      )
    ) {
      return ContentService
        .createTextOutput(
          JSON.stringify({
            success: false,
            code:
              'VIP_INACTIVE',
            message:
              'عضویت VIP شما فعال نیست.'
          })
        )
        .setMimeType(
          ContentService.MimeType.JSON
        );
    }
    const tokens =
      getVIPTokensForCustomer_(
        customer.id
      );
    return ContentService
      .createTextOutput(
        JSON.stringify({
          success: true,
          customer:
            customer,
          tokens:
            tokens
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );
  } catch (err) {
    return ContentService
      .createTextOutput(
        JSON.stringify({
          success: false,
          error:
            err.message
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );
  }
}
/* =========================================================
   VIP DATA
========================================================= */
function getVIPData(
  customerId
) {
  const resolvedCustomerId =
    extractCustomerId_(
      customerId
    );
  const customerRow =
    findCustomerRow_(
      resolvedCustomerId
    );
  if (
    customerRow === -1
  ) {
    throw new Error(
      'مشتری VIP پیدا نشد.'
    );
  }
  const customerSheet =
    getSheetOrThrow_(
      CUSTOMERS_SHEET
    );
  const customerValues =
    customerSheet
      .getRange(
        customerRow,
        1,
        1,
        8
      )
      .getValues()[0];
  const customer =
    getCustomerObject_(
      customerValues
    );
  const history = [];
  const payments = [];
  try {
    const historySheet =
      getSheetOrThrow_(
        HISTORY_SHEET
      );
    const lastRow =
      historySheet.getLastRow();
    if (lastRow >= 2) {
      const values =
        historySheet
          .getRange(
            2,
            1,
            lastRow - 1,
            historySheet.getLastColumn()
          )
          .getValues();
      values.forEach(
        function(row) {
          if (
            normalizeString_(
              row[0]
            ) ===
            resolvedCustomerId
          ) {
            history.push(
              row
            );
          }
        }
      );
    }
  } catch (err) {
  }
  try {
    const paymentSheet =
      getSheetOrThrow_(
        PAYMENTS_SHEET
      );
    const lastRow =
      paymentSheet.getLastRow();
    if (lastRow >= 2) {
      const values =
        paymentSheet
          .getRange(
            2,
            1,
            lastRow - 1,
            paymentSheet.getLastColumn()
          )
          .getValues();
      values.forEach(
        function(row) {
          if (
            normalizeString_(
              row[0]
            ) ===
            resolvedCustomerId
          ) {
            payments.push(
              row
            );
          }
        }
      );
    }
  } catch (err) {
  }
  const tokens =
    getVIPTokensForCustomer_(
      resolvedCustomerId
    );
  return {
    success: true,
    customer:
      customer,
    history:
      history,
    payments:
      payments,
    tokens:
      tokens
  };
}
/* =========================================================
   VIP TOKENS
========================================================= */
function getVIPTokensForCustomer_(
  customerId
) {
  const resolvedCustomerId =
    extractCustomerId_(
      customerId
    );
  const sheet =
    getSheetOrThrow_(
      TOKENS_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        Math.max(
          10,
          sheet.getLastColumn()
        )
      )
      .getValues();
  const now =
    new Date();
  const tokens = [];
  values.forEach(
    function(row) {
      if (
        normalizeString_(
          row[1]
        ) !==
        resolvedCustomerId
      ) {
        return;
      }
      let status =
        normalizeString_(
          row[5]
        );
      const expiryMs =
        Number(
          row[9] || 0
        );
      if (
        status !==
          'مصرف‌شده' &&
        expiryMs &&
        now.getTime() >=
          expiryMs
      ) {
        status =
          'منقضی‌شده';
      }
      tokens.push({
        code:
          normalizeString_(row[0]),
        customerId:
          normalizeString_(row[1]),
        discount:
          Number(row[2] || 0),
        discountPercent:
          Number(row[2] || 0),
        issuedAt:
          normalizeString_(row[3]),
        expiresAt:
          normalizeString_(row[4]),
        status:
          status,
        usedAt:
          normalizeString_(row[6]),
        trackingCode:
          normalizeString_(row[7])
      });
    }
  );
  return tokens;
}
function activateToken(
  customerId,
  code
) {
  const resolvedCustomerId =
    extractCustomerId_(
      customerId
    );
  const sheet =
    getSheetOrThrow_(
      TOKENS_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error(
      'توکنی وجود ندارد.'
    );
  }
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        Math.max(
          10,
          sheet.getLastColumn()
        )
      )
      .getValues();
  for (
    let i = 0;
    i < values.length;
    i++
  ) {
    const row =
      values[i];
    if (
      normalizeString_(
        row[0]
      ) ===
        normalizeString_(code) &&
      normalizeString_(
        row[1]
      ) ===
        resolvedCustomerId
    ) {
      if (
        normalizeString_(
          row[5]
        ) ===
        'مصرف‌شده'
      ) {
        throw new Error(
          'این توکن قبلاً مصرف شده است.'
        );
      }
      const expiryMs =
        Number(
          row[9] || 0
        );
      if (
        expiryMs &&
        Date.now() >=
          expiryMs
      ) {
        sheet
          .getRange(
            i + 2,
            6
          )
          .setValue(
            'منقضی‌شده'
          );
        throw new Error(
          'اعتبار این توکن به پایان رسیده است.'
        );
      }
      sheet
        .getRange(
          i + 2,
          6
        )
        .setValue(
          'فعال'
        );
      return {
        success: true,
        message:
          'توکن فعال شد.'
      };
    }
  }
  throw new Error(
    'توکن پیدا نشد.'
  );
}
function createVIPToken(
  customerId,
  discountPercent,
  requestId
) {
  const lock =
    LockService
      .getScriptLock();
  lock.waitLock(
    10000
  );
  try {
    const resolvedCustomerId =
      extractCustomerId_(
        customerId
      );
    const cache =
      CacheService
        .getScriptCache();
    if (requestId) {
      const cacheKey =
        'VIP_TOKEN_REQUEST_' +
        normalizeString_(
          requestId
        );
      const cached =
        cache.get(
          cacheKey
        );
      if (cached) {
        return JSON.parse(
          cached
        );
      }
    }
    const discount =
      Number(
        discountPercent
      );
    if (
      ALLOWED_DISCOUNTS.indexOf(
        discount
      ) === -1
    ) {
      throw new Error(
        'درصد تخفیف مجاز نیست.'
      );
    }
    const customerRow =
      findCustomerRow_(
        resolvedCustomerId
      );
    if (
      customerRow === -1
    ) {
      throw new Error(
        'مشتری VIP پیدا نشد.'
      );
    }
    const customerSheet =
      getSheetOrThrow_(
        CUSTOMERS_SHEET
      );
    const customer =
      customerSheet
        .getRange(
          customerRow,
          1,
          1,
          8
        )
        .getValues()[0];
    if (
      !isVIPActive_(
        customer[7]
      )
    ) {
      throw new Error(
        'عضویت این مشتری VIP فعال نیست.'
      );
    }
    const sheet =
      getSheetOrThrow_(
        TOKENS_SHEET
      );
    const now =
      new Date();
    const expiry =
      new Date(
        now.getTime() +
        TOKEN_EXPIRY_HOURS *
        60 *
        60 *
        1000
      );
    let code = '';
    let exists = true;
    while (exists) {
      code =
        generateRandomCode_(
          'VIP-',
          6
        );
      exists = false;
      const lastRow =
        sheet.getLastRow();
      if (lastRow >= 2) {
        const existing =
          sheet
            .getRange(
              2,
              1,
              lastRow - 1,
              1
            )
            .getValues();
        for (
          let i = 0;
          i < existing.length;
          i++
        ) {
          if (
            normalizeString_(
              existing[i][0]
            ) === code
          ) {
            exists = true;
            break;
          }
        }
      }
    }
    const result = {
      success: true,
      code:
        code,
      customerId:
        resolvedCustomerId,
      discount:
        discount,
      discountPercent:
        discount,
      issuedAt:
        formatPersianDateTime(
          now
        ),
      expiresAt:
        formatPersianDateTime(
          expiry
        ),
      status:
        'صادرشده',
      usedAt:
        '',
      trackingCode:
        ''
    };
    sheet.appendRow([
      code,
      resolvedCustomerId,
      discount,
      result.issuedAt,
      result.expiresAt,
      'صادرشده',
      '',
      '',
      now.getTime(),
      expiry.getTime()
    ]);
    if (requestId) {
      const cacheKey =
        'VIP_TOKEN_REQUEST_' +
        normalizeString_(
          requestId
        );
      cache.put(
        cacheKey,
        JSON.stringify(
          result
        ),
        60 * 10
      );
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}
function getVIPTokens() {
  const sheet =
    getSheetOrThrow_(
      TOKENS_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: true,
      tokens: []
    };
  }
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        Math.max(
          10,
          sheet.getLastColumn()
        )
      )
      .getValues();
  const now =
    Date.now();
  const tokens =
    values.map(
      function(row) {
        let status =
          normalizeString_(
            row[5]
          );
        const expiryMs =
          Number(
            row[9] || 0
          );
        if (
          status !==
            'مصرف‌شده' &&
          expiryMs &&
          now >= expiryMs
        ) {
          status =
            'منقضی‌شده';
        }
        return {
          code:
            normalizeString_(row[0]),
          customerId:
            normalizeString_(row[1]),
          discount:
            Number(row[2] || 0),
          discountPercent:
            Number(row[2] || 0),
          issuedAt:
            normalizeString_(row[3]),
          expiresAt:
            normalizeString_(row[4]),
          status:
            status,
          usedAt:
            normalizeString_(row[6]),
          trackingCode:
            normalizeString_(row[7])
        };
      }
    );
  return {
    success: true,
    tokens:
      tokens
  };
}
function validateVIPToken(
  code
) {
  const sheet =
    getSheetOrThrow_(
      TOKENS_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: false,
      valid: false,
      message:
        'توکن پیدا نشد.'
    };
  }
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        Math.max(
          10,
          sheet.getLastColumn()
        )
      )
      .getValues();
  for (
    let i = 0;
    i < values.length;
    i++
  ) {
    const row =
      values[i];
    if (
      normalizeString_(
        row[0]
      ) !==
      normalizeString_(
        code
      )
    ) {
      continue;
    }
    const status =
      normalizeString_(
        row[5]
      );
    const expiryMs =
      Number(
        row[9] || 0
      );
    if (
      status ===
      'مصرف‌شده'
    ) {
      return {
        success: true,
        valid: false,
        message:
          'این توکن قبلاً مصرف شده است.'
      };
    }
    if (
      expiryMs &&
      Date.now() >=
        expiryMs
    ) {
      sheet
        .getRange(
          i + 2,
          6
        )
        .setValue(
          'منقضی‌شده'
        );
      return {
        success: true,
        valid: false,
        message:
          'اعتبار این توکن به پایان رسیده است.'
      };
    }
    return {
      success: true,
      valid: true,
      code:
        normalizeString_(
          row[0]
        ),
      customerId:
        normalizeString_(
          row[1]
        ),
      discount:
        Number(row[2] || 0),
      discountPercent:
        Number(row[2] || 0),
      issuedAt:
        normalizeString_(
          row[3]
        ),
      expiresAt:
        normalizeString_(
          row[4]
        ),
      status:
        status,
      usedAt:
        normalizeString_(
          row[6]
        ),
      trackingCode:
        normalizeString_(
          row[7]
        )
    };
  }
  return {
    success: true,
    valid: false,
    message:
      'توکن پیدا نشد.'
  };
}
function useVIPToken(
  code,
  trackingCode
) {
  const lock =
    LockService
      .getScriptLock();
  lock.waitLock(
    10000
  );
  try {
    const sheet =
      getSheetOrThrow_(
        TOKENS_SHEET
      );
    const lastRow =
      sheet.getLastRow();
    if (lastRow < 2) {
      throw new Error(
        'توکن پیدا نشد.'
      );
    }
    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          Math.max(
            10,
            sheet.getLastColumn()
          )
        )
        .getValues();
    for (
      let i = 0;
      i < values.length;
      i++
    ) {
      const row =
        values[i];
      if (
        normalizeString_(
          row[0]
        ) !==
        normalizeString_(
          code
        )
      ) {
        continue;
      }
      if (
        normalizeString_(
          row[5]
        ) ===
        'مصرف‌شده'
      ) {
        throw new Error(
          'این توکن قبلاً مصرف شده است.'
        );
      }
      const expiryMs =
        Number(
          row[9] || 0
        );
      if (
        expiryMs &&
        Date.now() >=
          expiryMs
      ) {
        sheet
          .getRange(
            i + 2,
            6
          )
          .setValue(
            'منقضی‌شده'
          );
        throw new Error(
          'اعتبار این توکن به پایان رسیده است.'
        );
      }
      const usedAt =
        formatPersianDateTime(
          new Date()
        );
      sheet
        .getRange(
          i + 2,
          6
        )
        .setValue(
          'مصرف‌شده'
        );
      sheet
        .getRange(
          i + 2,
          7
        )
        .setValue(
          usedAt
        );
      if (trackingCode) {
        sheet
          .getRange(
            i + 2,
            8
          )
          .setValue(
            trackingCode
          );
      }
      return {
        success: true,
        code:
          code,
        customerId:
          normalizeString_(
            row[1]
          ),
        discount:
          Number(row[2] || 0),
        discountPercent:
          Number(row[2] || 0),
        usedAt:
          usedAt,
        trackingCode:
          trackingCode || ''
      };
    }
    throw new Error(
      'توکن پیدا نشد.'
    );
  } finally {
    lock.releaseLock();
  }
}
function calculateVIPDiscount(
  price,
  discountPercent
) {
  const originalPrice =
    Number(
      price || 0
    );
  const percent =
    Number(
      discountPercent || 0
    );
  const discountAmount =
    Math.round(
      originalPrice *
      percent /
      100
    );
  return {
    originalPrice:
      originalPrice,
    discountPercent:
      percent,
    discountAmount:
      discountAmount,
    finalPrice:
      originalPrice -
      discountAmount
  };
}
/* =========================================================
   VIP CUSTOMERS
========================================================= */
function getVIPCustomers() {
  const sheet =
    getSheetOrThrow_(
      CUSTOMERS_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: true,
      customers: []
    };
  }
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        8
      )
      .getValues();
  const customers =
    values.map(
      function(row) {
        return getCustomerObject_(
          row
        );
      }
    );
  return {
    success: true,
    customers:
      customers
  };
}
function addVIPCustomer(
  firstName,
  lastName,
  mobile,
  telegramId,
  activateNow
) {
  const sheet =
    getSheetOrThrow_(
      CUSTOMERS_SHEET
    );
  const lock =
    LockService
      .getScriptLock();
  lock.waitLock(
    10000
  );
  try {
    const cleanFirstName =
      normalizeString_(
        firstName
      );
    const cleanLastName =
      normalizeString_(
        lastName
      );
    const cleanMobile =
      normalizeString_(
        mobile
      );
    const cleanTelegramId =
      normalizeString_(
        telegramId
      );
    if (!cleanFirstName) {
      throw new Error(
        'نام را وارد کنید.'
      );
    }
    if (!cleanLastName) {
      throw new Error(
        'نام خانوادگی را وارد کنید.'
      );
    }
    const lastRow =
      sheet.getLastRow();
    let customerId =
      'VIP001';
    if (lastRow >= 2) {
      const ids =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            1
          )
          .getValues();
      let maxNumber =
        0;
      ids.forEach(
        function(row) {
          const id =
            normalizeString_(
              row[0]
            );
          const match =
            id.match(
              /^VIP(\d+)$/i
            );
          if (match) {
            maxNumber =
              Math.max(
                maxNumber,
                Number(
                  match[1]
                )
              );
          }
        }
      );
      customerId =
        'VIP' +
        String(
          maxNumber + 1
        ).padStart(
          3,
          '0'
        );
    }
    const joinedAt =
      formatPersianDateTime(
        new Date()
      );
    const status =
      activateNow === false
        ? 'غیرفعال'
        : 'فعال';
    sheet.appendRow([
      customerId,
      cleanFirstName,
      cleanLastName,
      cleanMobile,
      cleanTelegramId,
      joinedAt,
      0,
      status
    ]);
    return {
      success: true,
      customer: {
        id:
          customerId,
        firstName:
          cleanFirstName,
        lastName:
          cleanLastName,
        mobile:
          cleanMobile,
        telegramId:
          cleanTelegramId,
        joinedAt:
          joinedAt,
        bookingsCount:
          0,
        vipStatus:
          status
      }
    };
  } finally {
    lock.releaseLock();
  }
}
function setVIPStatus(
  customerId,
  status
) {
  const resolvedCustomerId =
    extractCustomerId_(
      customerId
    );
  const row =
    findCustomerRow_(
      resolvedCustomerId
    );
  if (row === -1) {
    throw new Error(
      'مشتری VIP پیدا نشد.'
    );
  }
  const cleanStatus =
    normalizeString_(
      status
    );
  if (
    cleanStatus !== 'فعال' &&
    cleanStatus !== 'غیرفعال'
  ) {
    throw new Error(
      'وضعیت VIP نامعتبر است.'
    );
  }
  const sheet =
    getSheetOrThrow_(
      CUSTOMERS_SHEET
    );
  sheet
    .getRange(
      row,
      8
    )
    .setValue(
      cleanStatus
    );
  return {
    success: true,
    customerId:
      resolvedCustomerId,
    vipStatus:
      cleanStatus
  };
}
/* =========================================================
   VIP CONNECTION CODES
========================================================= */
function getVIPConnectionCodes() {
  const sheet =
    getSheetOrThrow_(
      CONNECTION_CODES_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: true,
      codes: []
    };
  }
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        8
      )
      .getValues();
  const codes =
    values.map(
      function(row) {
        return {
          code:
            normalizeString_(
              row[0]
            ),
          customerId:
            normalizeString_(
              row[1]
            ),
          status:
            normalizeString_(
              row[2]
            ),
          createdAt:
            normalizeString_(
              row[3]
            ),
          usedAt:
            normalizeString_(
              row[4]
            ),
          telegramId:
            normalizeString_(
              row[5]
            ),
          createdSystemAt:
            row[6]
              ? String(row[6])
              : '',
          usedSystemAt:
            row[7]
              ? String(row[7])
              : '',
          connectionCode:
            normalizeString_(
              row[0]
            ),
          createdDate:
            normalizeString_(
              row[3]
            ),
          usedDate:
            normalizeString_(
              row[4]
            )
        };
      }
    );
  return {
    success: true,
    codes:
      codes
  };
}
/* =========================================================
   GENERATE VIP CONNECTION CODE
========================================================= */
function generateVIPConnectionCode(
  customerId
) {
  const lock =
    LockService
      .getScriptLock();
  lock.waitLock(
    10000
  );
  try {
    /*
     * تبدیل هر نوع ورودی Admin
     * به شناسه واقعی مشتری
     */
    const resolvedCustomerId =
      extractCustomerId_(
        customerId
      );
    if (!resolvedCustomerId) {
      throw new Error(
        'شناسه مشتری VIP مشخص نیست.'
      );
    }
    const customerRow =
      findCustomerRow_(
        resolvedCustomerId
      );
    if (
      customerRow === -1
    ) {
      throw new Error(
        'مشتری VIP پیدا نشد.'
      );
    }
    const customerSheet =
      getSheetOrThrow_(
        CUSTOMERS_SHEET
      );
    const customer =
      customerSheet
        .getRange(
          customerRow,
          1,
          1,
          8
        )
        .getValues()[0];
    if (
      !isVIPActive_(
        customer[7]
      )
    ) {
      throw new Error(
        'عضویت این مشتری VIP فعال نیست.'
      );
    }
    const sheet =
      getSheetOrThrow_(
        CONNECTION_CODES_SHEET
      );
    const lastRow =
      sheet.getLastRow();
    let existing = [];
    if (lastRow >= 2) {
      existing =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            8
          )
          .getValues();
    }
    /*
     * اگر برای این مشتری قبلاً
     * کد فعال وجود داشته باشد،
     * همان کد را برمی‌گردانیم.
     */
    for (
      let i = 0;
      i < existing.length;
      i++
    ) {
      const row =
        existing[i];
      if (
        normalizeString_(
          row[1]
        ) ===
        resolvedCustomerId &&
        normalizeString_(
          row[2]
        ) ===
        CONNECTION_CODE_ACTIVE_STATUS
      ) {
        return {
          success: true,
          code:
            normalizeString_(
              row[0]
            ),
          customerId:
            normalizeString_(
              row[1]
            ),
          status:
            normalizeString_(
              row[2]
            ),
          createdAt:
            normalizeString_(
              row[3]
            ),
          usedAt:
            normalizeString_(
              row[4]
            ),
          telegramId:
            normalizeString_(
              row[5]
            ),
          createdSystemAt:
            row[6]
              ? String(row[6])
              : '',
          usedSystemAt:
            row[7]
              ? String(row[7])
              : '',
          existing:
            true
        };
      }
    }
    let code = '';
    let exists = true;
    while (exists) {
      code =
        generateRandomCode_(
          'VIP-CON-',
          6
        );
      exists = false;
      for (
        let i = 0;
        i < existing.length;
        i++
      ) {
        if (
          normalizeString_(
            existing[i][0]
          ) === code
        ) {
          exists = true;
          break;
        }
      }
    }
    const now =
      new Date();
    const createdAt =
      formatPersianDateTime(
        now
      );
    sheet.appendRow([
      code,
      resolvedCustomerId,
      CONNECTION_CODE_ACTIVE_STATUS,
      createdAt,
      '',
      '',
      now.getTime(),
      ''
    ]);
    return {
      success: true,
      code:
        code,
      customerId:
        resolvedCustomerId,
      status:
        CONNECTION_CODE_ACTIVE_STATUS,
      createdAt:
        createdAt,
      usedAt:
        '',
      telegramId:
        '',
      createdSystemAt:
        String(
          now.getTime()
        ),
      usedSystemAt:
        '',
      existing:
        false
    };
  } finally {
    lock.releaseLock();
  }
}
/* =========================================================
   REVOKE CONNECTION CODE
========================================================= */
function revokeVIPConnectionCode(
  code
) {
  const cleanCode =
    normalizeString_(
      code
    );
  if (!cleanCode) {
    throw new Error(
      'کد اتصال مشخص نشده است.'
    );
  }
  const sheet =
    getSheetOrThrow_(
      CONNECTION_CODES_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error(
      'کد اتصالی وجود ندارد.'
    );
  }
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        8
      )
      .getValues();
  for (
    let i = 0;
    i < values.length;
    i++
  ) {
    if (
      normalizeString_(
        values[i][0]
      ) === cleanCode
    ) {
      const currentStatus =
        normalizeString_(
          values[i][2]
        );
      if (
        currentStatus ===
        CONNECTION_CODE_USED_STATUS
      ) {
        throw new Error(
          'کد استفاده شده را نمی‌توان باطل کرد.'
        );
      }
      sheet
        .getRange(
          i + 2,
          3
        )
        .setValue(
          CONNECTION_CODE_REVOKED_STATUS
        );
      return {
        success: true,
        code:
          cleanCode,
        status:
          CONNECTION_CODE_REVOKED_STATUS
      };
    }
  }
  throw new Error(
    'کد اتصال پیدا نشد.'
  );
}
/* =========================================================
   CONNECT TELEGRAM TO VIP
========================================================= */
function connectTelegramToVIP(
  code,
  telegramId
) {
  const cleanCode =
    normalizeString_(
      code
    );
  const cleanTelegramId =
    normalizeString_(
      telegramId
    );
  if (!cleanCode) {
    throw new Error(
      'کد اتصال وارد نشده است.'
    );
  }
  if (!cleanTelegramId) {
    throw new Error(
      'Telegram ID وارد نشده است.'
    );
  }
  const lock =
    LockService
      .getScriptLock();
  lock.waitLock(
    10000
  );
  try {
    const sheet =
      getSheetOrThrow_(
        CONNECTION_CODES_SHEET
      );
    const lastRow =
      sheet.getLastRow();
    if (lastRow < 2) {
      throw new Error(
        'کد اتصال پیدا نشد.'
      );
    }
    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          8
        )
        .getValues();
    let targetIndex =
      -1;
    let targetRow =
      null;
    for (
      let i = 0;
      i < values.length;
      i++
    ) {
      if (
        normalizeString_(
          values[i][0]
        ) === cleanCode
      ) {
        targetIndex =
          i;
        targetRow =
          values[i];
        break;
      }
    }
    if (
      targetIndex === -1
    ) {
      throw new Error(
        'کد اتصال پیدا نشد.'
      );
    }
    const status =
      normalizeString_(
        targetRow[2]
      );
    if (
      status !==
      CONNECTION_CODE_ACTIVE_STATUS
    ) {
      if (
        status ===
        CONNECTION_CODE_USED_STATUS
      ) {
        throw new Error(
          'این کد قبلاً استفاده شده است.'
        );
      }
      throw new Error(
        'این کد اتصال دیگر معتبر نیست.'
      );
    }
    const customerId =
      normalizeString_(
        targetRow[1]
      );
    const customerRow =
      findCustomerRow_(
        customerId
      );
    if (
      customerRow === -1
    ) {
      throw new Error(
        'مشتری مربوط به این کد پیدا نشد.'
      );
    }
    const customerSheet =
      getSheetOrThrow_(
        CUSTOMERS_SHEET
      );
    const customer =
      customerSheet
        .getRange(
          customerRow,
          1,
          1,
          8
        )
        .getValues()[0];
    if (
      !isVIPActive_(
        customer[7]
      )
    ) {
      throw new Error(
        'عضویت VIP این مشتری فعال نیست.'
      );
    }
    const existingTelegramId =
      normalizeString_(
        customer[4]
      );
    if (
      existingTelegramId &&
      existingTelegramId !==
      cleanTelegramId
    ) {
      throw new Error(
        'این مشتری قبلاً به یک حساب Telegram دیگر متصل شده است.'
      );
    }
    const now =
      new Date();
    const usedAt =
      formatPersianDateTime(
        now
      );
    customerSheet
      .getRange(
        customerRow,
        5
      )
      .setValue(
        cleanTelegramId
      );
    sheet
      .getRange(
        targetIndex + 2,
        3
      )
      .setValue(
        CONNECTION_CODE_USED_STATUS
      );
    sheet
      .getRange(
        targetIndex + 2,
        5
      )
      .setValue(
        usedAt
      );
    sheet
      .getRange(
        targetIndex + 2,
        6
      )
      .setValue(
        cleanTelegramId
      );
    sheet
      .getRange(
        targetIndex + 2,
        8
      )
      .setValue(
        now.getTime()
      );
    return {
      success: true,
      customer:
        getCustomerObject_(
          customerSheet
            .getRange(
              customerRow,
              1,
              1,
              8
            )
            .getValues()[0]
        ),
      code:
        cleanCode,
      status:
        CONNECTION_CODE_USED_STATUS,
      usedAt:
        usedAt
    };
  } finally {
    lock.releaseLock();
  }
}
/* =========================================================
   VIP SETTINGS
========================================================= */
function getVIPSettings() {
  return {
    success: true,
    settings: {
      bookingUrl:
        'https://kaenatchi.github.io/booking/'
    }
  };
}
function saveVIPSettings(
  settings
) {
  const data =
    settings || {};
  return {
    success: true,
    settings: {
      bookingUrl:
        normalizeString_(
          data.bookingUrl ||
          'https://kaenatchi.github.io/booking/'
        )
    }
  };
}
/* =========================================================
   DATE HELPERS — JALALI
========================================================= */
function formatPersianDate(
  date
) {
  if (!date) {
    return '';
  }
  const d =
    date instanceof Date
      ? date
      : new Date(date);
  const parts =
    gregorianToPersianDate_(
      d
    );
  return (
    parts[0] +
    '/' +
    String(
      parts[1]
    ).padStart(
      2,
      '0'
    ) +
    '/' +
    String(
      parts[2]
    ).padStart(
      2,
      '0'
    )
  );
}
function formatPersianDateTime(
  date
) {
  if (!date) {
    return '';
  }
  const d =
    date instanceof Date
      ? date
      : new Date(date);
  const datePart =
    formatPersianDate(
      d
    );
  const hours =
    String(
      d.getHours()
    ).padStart(
      2,
      '0'
    );
  const minutes =
    String(
      d.getMinutes()
    ).padStart(
      2,
      '0'
    );
  return (
    datePart +
    ' - ' +
    hours +
    ':' +
    minutes
  );
}
function gregorianToPersianDate_(
  date
) {
  let gy =
    date.getFullYear();
  const gm =
    date.getMonth() + 1;
  const gd =
    date.getDate();
  const gDaysInMonth = [
    31,
    28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31
  ];
  const gy2 =
    gm > 2
      ? gy + 1
      : gy;
  let days =
    355666 +
    365 * gy +
    Math.floor(
      (gy2 + 3) / 4
    ) -
    Math.floor(
      (gy2 + 99) / 100
    ) +
    Math.floor(
      (gy2 + 399) / 400
    );
  for (
    let i = 0;
    i < gm - 1;
    i++
  ) {
    days +=
      gDaysInMonth[i];
  }
  if (
    gm > 2 &&
    gy % 4 === 0 &&
    (
      gy % 100 !== 0 ||
      gy % 400 === 0
    )
  ) {
    days++;
  }
  days += gd;
  let jy =
    -1595 +
    33 *
    Math.floor(
      days / 12053
    );
  days %= 12053;
  jy +=
    4 *
    Math.floor(
      days / 1461
    );
  days %= 1461;
  if (days > 365) {
    jy +=
      Math.floor(
        (days - 1) / 365
      );
    days =
      (days - 1) % 365;
  }
  let jm;
  let jd;
  if (days < 186) {
    jm =
      1 +
      Math.floor(
        days / 31
      );
    jd =
      1 +
      days % 31;
  } else {
    jm =
      7 +
      Math.floor(
        (days - 186) / 30
      );
    jd =
      1 +
      (
        (days - 186) % 30
      );
  }
  return [
    jy,
    jm,
    jd
  ];
}
/* =========================================================
   TEST FUNCTIONS
========================================================= */
function testVIPCustomer() {
  const result =
    getVIPCustomers();
  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
  return result;
}
function testWebApp() {
  return {
    success: true,
    message:
      'Web App backend is running.'
  };
}
function testCreateVIPToken() {
  const customers =
    getVIPCustomers();
  if (
    !customers.customers ||
    !customers.customers.length
  ) {
    throw new Error(
      'مشتری VIP وجود ندارد.'
    );
  }
  return createVIPToken(
    customers
      .customers[0]
      .id,
    3,
    'TEST-' +
    Date.now()
  );
}
function testGetVIPCustomers() {
  return getVIPCustomers();
}
function testAdminFile() {
  return HtmlService
    .createHtmlOutputFromFile(
      'Admin'
    )
    .getContent()
    .substring(
      0,
      500
    );
}
function testGetVIPTokens() {
  return getVIPTokens();
}
function testTelegramVIP() {
  return {
    success: true,
    message:
      'برای تست واقعی Telegram باید initData معتبر ارسال شود.'
  };
}
function testGetVIPCustomersResult() {
  const result =
    getVIPCustomers();
  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
  return result;
}
function testCreateVIPTokenResult() {
  return testCreateVIPToken();
}
function testVIPJoinedAtRaw() {
  const sheet =
    getSheetOrThrow_(
      CUSTOMERS_SHEET
    );
  const lastRow =
    sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  const value =
    sheet
      .getRange(
        2,
        6
      )
      .getValue();
  Logger.log(
    JSON.stringify(
      value
    )
  );
  return value;
}
function testGetVIPConnectionCodes() {
  const result =
    getVIPConnectionCodes();
  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
  return result;
}
function testConnectTelegram() {
  return {
    success: true,
    message:
      'برای اتصال واقعی باید کد اتصال و Telegram ID معتبر ارسال شود.'
  };
}

function testVIPCustomerLookup() {
  const sheet = getSheetOrThrow_(CUSTOMERS_SHEET);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log('CUSTOMERS SHEET IS EMPTY');
    return;
  }

  const rows = sheet
    .getRange(2, 1, lastRow - 1, 8)
    .getValues();

  Logger.log('--- VIP CUSTOMERS ---');

  for (let i = 0; i < rows.length; i++) {
    Logger.log(
      'ROW ' + (i + 2) +
      ' | ID=[' + String(rows[i][0]) + ']' +
      ' | FIRST=[' + String(rows[i][1]) + ']' +
      ' | LAST=[' + String(rows[i][2]) + ']' +
      ' | VIP=[' + String(rows[i][7]) + ']'
    );
  }

  Logger.log('--- END ---');
}

function testFindVIP001() {
  const customerId = 'VIP001';

  const resolved =
    extractCustomerId_(customerId);

  const row =
    findCustomerRow_(customerId);

  Logger.log(
    'INPUT = [' +
    customerId +
    ']'
  );

  Logger.log(
    'RESOLVED = [' +
    resolved +
    ']'
  );

  Logger.log(
    'ROW = [' +
    row +
    ']'
  );
}

function testGenerateVIPConnectionCode() {
  const result =
    generateVIPConnectionCode('VIP001');

  Logger.log(
    JSON.stringify(result)
  );
}
