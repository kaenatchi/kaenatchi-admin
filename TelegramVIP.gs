/*******************************************************
 * KaenatChi VIP Telegram Mini App Bridge
 *******************************************************/

const TELEGRAM_VIP_SPREADSHEET_ID =
  "1TpljGwyRpHcxyyR6Zm1mbT2CPJBvFeHCvD3QH34BuGs";

const TELEGRAM_VIP_CUSTOMERS_SHEET =
  "مشتریان";

const TELEGRAM_VIP_BOOKINGS_SHEET =
  "سوابق نوبت ها";

const TELEGRAM_VIP_PAYMENTS_SHEET =
  "پرداخت ها";

const TELEGRAM_VIP_TOKENS_SHEET =
  "توکن ها";


/**
 * JSON response helper
 */
function telegramVipJson_(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


/**
 * Telegram Mini App POST endpoint
 *
 * Supported actions:
 *
 * load
 * connect
 */
function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {

      return telegramVipJson_({
        success: false,
        error: "INVALID_REQUEST"
      });
    }

    const body =
      JSON.parse(
        e.postData.contents
      );

    const initData =
      String(body.initData || '').trim();

    const action =
      String(body.action || 'load')
        .trim()
        .toLowerCase();

    if (!initData) {

      return telegramVipJson_({
        success: false,
        error: "MISSING_INIT_DATA"
      });
    }

    /*
     * Validate Telegram initData
     */
    const validation =
      validateTelegramInitData_(initData);

    if (!validation.valid) {

      return telegramVipJson_({
        success: false,
        error: "INVALID_TELEGRAM_AUTH"
      });
    }

    const telegramId =
      String(
        validation.telegramId || ''
      ).trim();

    if (!telegramId) {

      return telegramVipJson_({
        success: false,
        error: "TELEGRAM_ID_MISSING"
      });
    }


    /*
     * First-time connection
     */
    if (action === 'connect') {

      const connectionCode =
        String(
          body.connectionCode || ''
        ).trim();

      if (!connectionCode) {

        return telegramVipJson_({
          success: false,
          error: "MISSING_CONNECTION_CODE"
        });
      }

      const connectionResult =
        connectVIPTelegramAccount_(
          telegramId,
          connectionCode
        );

      if (!connectionResult.success) {

        return telegramVipJson_({
          success: false,
          error: "CONNECTION_FAILED"
        });
      }

      /*
       * بعد از اتصال موفق، داشبورد را برمی‌گردانیم.
       */
      return buildTelegramVipDashboardResponse_(
        connectionResult.customerId
      );
    }


    /*
     * Default action = load
     */
    if (action === 'load') {

      const customer =
        findTelegramCustomer_(
          telegramId
        );

      /*
       * Telegram ID هنوز به هیچ مشتری VIP متصل نیست.
       *
       * مهم:
       * Telegram ID را به کلاینت برنمی‌گردانیم.
       */
      if (!customer) {

        return telegramVipJson_({
          success: false,
          needsConnectionCode: true,
          error: "NEEDS_CONNECTION_CODE"
        });
      }

      /*
       * VIP غیرفعال
       */
      if (
        String(customer.vipStatus || '').trim() !==
        'فعال'
      ) {

        return telegramVipJson_({
          success: false,
          accessDenied: true,
          error: "VIP_INACTIVE"
        });
      }

      return buildTelegramVipDashboardResponse_(
        customer.customerId
      );
    }


    return telegramVipJson_({
      success: false,
      error: "UNKNOWN_ACTION"
    });

  } catch (error) {

    console.error(
      'Telegram VIP error:',
      error
    );

    return telegramVipJson_({
      success: false,
      error: "SERVER_ERROR"
    });
  }
}


/**
 * Telegram initData validation
 */
function validateTelegramInitData_(initData) {

  try {

    const botToken =
      PropertiesService
        .getScriptProperties()
        .getProperty("TELEGRAM_BOT_TOKEN");

    if (!botToken) {

      console.error(
        "TELEGRAM_BOT_TOKEN is missing."
      );

      return {
        valid: false
      };
    }

    const params = {};

    initData
      .split('&')
      .forEach(function(part) {

        if (!part) {
          return;
        }

        const separator =
          part.indexOf('=');

        if (separator === -1) {
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

        params[key] = value;
      });


    const receivedHash =
      String(params.hash || '')
        .toLowerCase();

    if (!receivedHash) {

      return {
        valid: false
      };
    }


    const dataCheckString =
      Object.keys(params)
        .filter(
          key => key !== 'hash'
        )
        .sort()
        .map(
          key =>
            key + '=' + params[key]
        )
        .join('\n');


    const botTokenBytes =
      Utilities
        .newBlob(botToken)
        .getBytes();

    const webAppDataKeyBytes =
      Utilities
        .newBlob("WebAppData")
        .getBytes();


    const secretKey =
      Utilities
        .computeHmacSha256Signature(
          botTokenBytes,
          webAppDataKeyBytes
        );


    const dataCheckStringBytes =
      Utilities
        .newBlob(dataCheckString)
        .getBytes();


    const calculatedHashBytes =
      Utilities
        .computeHmacSha256Signature(
          dataCheckStringBytes,
          secretKey
        );


    const calculatedHash =
      bytesToHex_(
        calculatedHashBytes
      ).toLowerCase();


    if (calculatedHash !== receivedHash) {

      return {
        valid: false
      };
    }


    /*
     * auth_date validation
     */
    const authDate =
      Number(params.auth_date || 0);

    if (!authDate) {

      return {
        valid: false
      };
    }


    const nowSeconds =
      Math.floor(
        Date.now() / 1000
      );


    /*
     * جلوگیری از auth_date آینده
     */
    if (
      authDate >
      nowSeconds + 300
    ) {

      return {
        valid: false
      };
    }


    /*
     * حداکثر عمر 24 ساعت
     */
    if (
      nowSeconds - authDate >
      86400
    ) {

      return {
        valid: false
      };
    }


    /*
     * استخراج Telegram ID
     */
    let telegramId = '';

    if (params.user) {

      try {

        const user =
          JSON.parse(
            params.user
          );

        telegramId =
          String(
            user.id || ''
          );

      } catch (error) {

        return {
          valid: false
        };
      }
    }


    if (!telegramId) {

      return {
        valid: false
      };
    }


    return {
      valid: true,
      telegramId: telegramId
    };

  } catch (error) {

    console.error(
      'Telegram initData validation failed:',
      error
    );

    return {
      valid: false
    };
  }
}


/**
 * تبدیل Byte[] به Hex
 */
function bytesToHex_(bytes) {

  return bytes
    .map(function(byte) {

      const value =
        byte < 0
          ? byte + 256
          : byte;

      return (
        '0' +
        value.toString(16)
      ).slice(-2);

    })
    .join('');
}


/**
 * پیدا کردن مشتری بر اساس Telegram ID
 *
 * Telegram ID در ستون E است.
 */
function findTelegramCustomer_(telegramId) {

  const ss =
    SpreadsheetApp.openById(
      TELEGRAM_VIP_SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      TELEGRAM_VIP_CUSTOMERS_SHEET
    );

  if (!sheet) {
    throw new Error(
      "Customers sheet not found."
    );
  }

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
    String(
      telegramId || ''
    ).trim();


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const rowTelegramId =
      String(
        values[i][4] || ''
      ).trim();

    if (
      rowTelegramId === target
    ) {

      return {
        customerId:
          String(
            values[i][0] || ''
          ).trim(),

        firstName:
          String(
            values[i][1] || ''
          ),

        lastName:
          String(
            values[i][2] || ''
          ),

        mobile:
          String(
            values[i][3] || ''
          ),

        telegramId:
          rowTelegramId,

        joinedAt:
          values[i][5],

        bookingCount:
          values[i][6],

        vipStatus:
          String(
            values[i][7] || ''
          ).trim()
      };
    }
  }

  return null;
}


/**
 * ساخت پاسخ کامل داشبورد VIP
 */
function buildTelegramVipDashboardResponse_(
  customerId
) {

  const ss =
    SpreadsheetApp.openById(
      TELEGRAM_VIP_SPREADSHEET_ID
    );

  const customer =
    findCustomerForTelegramDashboard_(
      ss,
      customerId
    );

  if (!customer) {

    return telegramVipJson_({
      success: false,
      error: "CUSTOMER_NOT_FOUND"
    });
  }


  if (
    String(customer.vipStatus || '').trim() !==
    'فعال'
  ) {

    return telegramVipJson_({
      success: false,
      accessDenied: true,
      error: "VIP_INACTIVE"
    });
  }


  const bookings =
    getTelegramVipBookings_(
      ss,
      customerId
    );

  const payments =
    getTelegramVipPayments_(
      ss,
      customerId
    );

  const tokens =
    getTelegramVipTokens_(
      ss,
      customerId
    );


  return telegramVipJson_({

    success: true,

    customer: customer,

    bookings: bookings,

    payments: payments,

    tokens: tokens

  });
}


/**
 * گرفتن اطلاعات مشتری برای داشبورد
 */
function findCustomerForTelegramDashboard_(
  ss,
  customerId
) {

  const sheet =
    ss.getSheetByName(
      TELEGRAM_VIP_CUSTOMERS_SHEET
    );

  if (!sheet) {
    throw new Error(
      "Customers sheet not found."
    );
  }

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
    String(
      customerId || ''
    ).trim();


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const id =
      String(
        values[i][0] || ''
      ).trim();

    if (id !== target) {
      continue;
    }


    return {

      customerId: id,

      firstName:
        String(
          values[i][1] || ''
        ),

      lastName:
        String(
          values[i][2] || ''
        ),

      mobile:
        String(
          values[i][3] || ''
        ),

      telegramId:
        String(
          values[i][4] || ''
        ),

      joinedAt:
        formatTelegramDate_(
          values[i][5]
        ),

      bookingCount:
        Number(
          values[i][6] || 0
        ),

      vipStatus:
        String(
          values[i][7] || ''
        )
    };
  }

  return null;
}


/**
 * فرمت تاریخ
 */
function formatTelegramDate_(value) {

  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {

    if (
      typeof formatPersianDateTime ===
      'function'
    ) {

      return formatPersianDateTime(
        value
      );
    }

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy/MM/dd HH:mm:ss'
    );
  }

  return String(value || '');
}


/**
 * سوابق نوبت
 */
function getTelegramVipBookings_(
  ss,
  customerId
) {

  const sheet =
    ss.getSheetByName(
      TELEGRAM_VIP_BOOKINGS_SHEET
    );

  if (!sheet) {
    return [];
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {

    return [];
  }


  const values =
    sheet
      .getRange(
        1,
        1,
        lastRow,
        lastColumn
      )
      .getValues();


  const headers =
    values[0];

  const target =
    String(
      customerId || ''
    ).trim();


  const result = [];


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const row =
      values[i];

    const rowText =
      row
        .map(
          value =>
            String(value || '')
        )
        .join(' ');


    if (
      rowText.indexOf(target) !== -1
    ) {

      result.push(
        rowToObject_(
          headers,
          row
        )
      );
    }
  }


  return result.reverse();
}


/**
 * پرداخت‌ها
 */
function getTelegramVipPayments_(
  ss,
  customerId
) {

  const sheet =
    ss.getSheetByName(
      TELEGRAM_VIP_PAYMENTS_SHEET
    );

  if (!sheet) {
    return [];
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {

    return [];
  }


  const values =
    sheet
      .getRange(
        1,
        1,
        lastRow,
        lastColumn
      )
      .getValues();


  const headers =
    values[0];

  const target =
    String(
      customerId || ''
    ).trim();


  const result = [];


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const row =
      values[i];

    const rowText =
      row
        .map(
          value =>
            String(value || '')
        )
        .join(' ');


    if (
      rowText.indexOf(target) !== -1
    ) {

      result.push(
        rowToObject_(
          headers,
          row
        )
      );
    }
  }


  return result.reverse();
}


/**
 * توکن‌ها
 */
function getTelegramVipTokens_(
  ss,
  customerId
) {

  const sheet =
    ss.getSheetByName(
      TELEGRAM_VIP_TOKENS_SHEET
    );

  if (!sheet) {
    return [];
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {

    return [];
  }


  const values =
    sheet
      .getRange(
        1,
        1,
        lastRow,
        lastColumn
      )
      .getValues();


  const headers =
    values[0];

  const target =
    String(
      customerId || ''
    ).trim();


  const result = [];


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const row =
      values[i];

    const rowText =
      row
        .map(
          value =>
            String(value || '')
        )
        .join(' ');


    if (
      rowText.indexOf(target) !== -1
    ) {

      result.push(
        rowToObject_(
          headers,
          row
        )
      );
    }
  }


  return result.reverse();
}


/**
 * تبدیل ردیف به Object
 */
function rowToObject_(
  headers,
  row
) {

  const object = {};

  for (
    let i = 0;
    i < headers.length;
    i++
  ) {

    const key =
      String(
        headers[i] || ''
      ).trim();

    if (!key) {
      continue;
    }


    let value =
      row[i];


    if (
      value instanceof Date &&
      !isNaN(value.getTime())
    ) {

      value =
        Utilities.formatDate(
          value,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd HH:mm:ss'
        );
    }


    object[key] =
      value;
  }


  return object;
}
