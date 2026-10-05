/*******************************************************
 * KaenatChi VIP - Connection Code System
 * Secure one-time Telegram account connection
 *******************************************************/

const VIP_CONNECTION_SHEET = 'کدهای اتصال VIP';
const VIP_CONNECTION_CODE_LENGTH = 6;


/**
 * دریافت شیت کدهای اتصال
 * اگر شیت وجود نداشته باشد، خودکار ساخته می‌شود.
 */
function getVIPConnectionSheet_() {

  const ss =
    SpreadsheetApp.openById(
      VIP_SHEET_ID
    );

  let sheet =
    ss.getSheetByName(
      VIP_CONNECTION_SHEET
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        VIP_CONNECTION_SHEET
      );

    sheet
      .getRange(
        1,
        1,
        1,
        8
      )
      .setValues([[
        'کد اتصال',
        'شناسه مشتری',
        'وضعیت',
        'تاریخ ایجاد',
        'تاریخ استفاده',
        'Telegram ID',
        'زمان سیستم ایجاد',
        'زمان سیستم استفاده'
      ]]);

    sheet.setFrozenRows(1);
  }

  return sheet;
}


/**
 * نرمال‌سازی کد اتصال
 */
function normalizeVIPConnectionCode_(code) {

  return String(
    code || ''
  )
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '');
}


/**
 * نرمال‌سازی شناسه مشتری
 */
function normalizeVIPCustomerId_(customerId) {

  return String(
    customerId || ''
  )
    .trim();
}


/**
 * نرمال‌سازی Telegram ID
 */
function normalizeVIPTelegramId_(telegramId) {

  return String(
    telegramId || ''
  )
    .trim();
}


/**
 * تبدیل نتیجه findCustomerRow_ به شماره واقعی ردیف
 *
 * برای سازگاری با نسخه‌های مختلف helper
 * هر دو حالت Number و Object را پشتیبانی می‌کند.
 */
function getVIPCustomerRowNumber_(result) {

  if (
    result === null ||
    result === undefined
  ) {

    return -1;
  }

  if (
    typeof result === 'number'
  ) {

    return result > 0
      ? result
      : -1;
  }

  if (
    typeof result === 'object'
  ) {

    if (
      result.rowNumber !== undefined &&
      result.rowNumber !== null
    ) {

      const rowNumber =
        Number(
          result.rowNumber
        );

      return (
        Number.isFinite(
          rowNumber
        ) &&
        rowNumber > 0
      )
        ? rowNumber
        : -1;
    }

    if (
      result.row !== undefined &&
      result.row !== null
    ) {

      const rowNumber =
        Number(
          result.row
        );

      return (
        Number.isFinite(
          rowNumber
        ) &&
        rowNumber > 0
      )
        ? rowNumber
        : -1;
    }
  }

  return -1;
}


/**
 * ساخت کد تصادفی
 */
function generateRandomVIPConnectionCode_() {

  const chars =
    'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

  let result = '';

  for (
    let i = 0;
    i < VIP_CONNECTION_CODE_LENGTH;
    i++
  ) {

    const index =
      Math.floor(
        Math.random() *
        chars.length
      );

    result +=
      chars.charAt(
        index
      );
  }

  return (
    'VIP-CON-' +
    result
  );
}


/**
 * پیدا کردن ردیف کد اتصال
 */
function findVIPConnectionCodeRow_(
  sheet,
  code
) {

  const cleanCode =
    normalizeVIPConnectionCode_(
      code
    );

  if (!cleanCode) {
    return -1;
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return -1;
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

    const rowCode =
      normalizeVIPConnectionCode_(
        values[i][0]
      );

    if (
      rowCode ===
      cleanCode
    ) {

      return i + 2;
    }
  }

  return -1;
}


/**
 * ساخت کد اتصال جدید برای مشتری
 *
 * هر مشتری فقط یک کد صادرشده دارد.
 * با ساخت کد جدید، کد قبلی همان مشتری باطل می‌شود.
 */
function generateVIPConnectionCode(
  customerId
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(15000);

  try {

    const cleanCustomerId =
      normalizeVIPCustomerId_(
        customerId
      );

    if (!cleanCustomerId) {

      throw new Error(
        'شناسه مشتری وارد نشده است.'
      );
    }


    const customerSheet =
      getSheetOrThrow_(
        CUSTOMERS_SHEET
      );


    /*
     * findCustomerRow_ فقط شناسه مشتری
     * را دریافت می‌کند و شماره ردیف را برمی‌گرداند.
     */
    const customerRow =
      findCustomerRow_(
        cleanCustomerId
      );


    if (
      customerRow === -1
    ) {

      throw new Error(
        'مشتری پیدا نشد.'
      );
    }


    const customerData =
      customerSheet
        .getRange(
          customerRow,
          1,
          1,
          8
        )
        .getValues()[0];


    const actualCustomerId =
      normalizeVIPCustomerId_(
        customerData[0]
      );


    if (
      actualCustomerId !==
      cleanCustomerId
    ) {

      throw new Error(
        'شناسه مشتری با اطلاعات شیت مطابقت ندارد.'
      );
    }


    const vipStatus =
      customerData[7];


    if (
      typeof isVIPActive_ ===
        'function' &&
      !isVIPActive_(
        vipStatus
      )
    ) {

      throw new Error(
        'برای مشتری غیرفعال نمی‌توان کد اتصال ساخت.'
      );
    }


    const connectionSheet =
      getVIPConnectionSheet_();


    const lastRow =
      connectionSheet.getLastRow();


    /*
     * کد قبلی همین مشتری را باطل می‌کنیم.
     */
    if (
      lastRow >= 2
    ) {

      const values =
        connectionSheet
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

        const rowCustomerId =
          normalizeVIPCustomerId_(
            values[i][1]
          );


        const rowStatus =
          String(
            values[i][2] || ''
          ).trim();


        if (
          rowCustomerId ===
            cleanCustomerId &&
          rowStatus ===
            'صادرشده'
        ) {

          connectionSheet
            .getRange(
              i + 2,
              3
            )
            .setValue(
              'باطل‌شده'
            );
        }
      }
    }


    /*
     * تولید کد یکتا
     */
    let code = '';

    let attempts = 0;


    do {

      code =
        generateRandomVIPConnectionCode_();

      attempts++;


      if (
        attempts > 100
      ) {

        throw new Error(
          'امکان ساخت کد یکتا وجود ندارد. دوباره تلاش کنید.'
        );
      }

    } while (
      findVIPConnectionCodeRow_(
        connectionSheet,
        code
      ) !== -1
    );


    const now =
      new Date();


    const displayDate =
      typeof formatPersianDateTime ===
        'function'

        ? formatPersianDateTime(
            now
          )

        : Utilities.formatDate(
            now,
            Session.getScriptTimeZone(),
            'yyyy/MM/dd HH:mm:ss'
          );


    connectionSheet.appendRow([
      code,
      cleanCustomerId,
      'صادرشده',
      displayDate,
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
        cleanCustomerId,

      status:
        'صادرشده',

      issuedAt:
        displayDate
    };


  } finally {

    lock.releaseLock();
  }
}


/**
 * دریافت تمام کدهای اتصال برای پنل مدیریت
 */
function getVIPConnectionCodes() {

  const sheet =
    getVIPConnectionSheet_();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return [];
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


  return values

    .filter(
      function(row) {

        return String(
          row[0] || ''
        ).trim() !== '';

      }
    )

    .map(
      function(row) {

        return {

          code:
            String(
              row[0] || ''
            ),

          customerId:
            String(
              row[1] || ''
            ),

          status:
            String(
              row[2] || ''
            ),

          issuedAt:
            String(
              row[3] || ''
            ),

          usedAt:
            String(
              row[4] || ''
            ),

          connected:
            Boolean(
              String(
                row[5] || ''
              ).trim()
            )
        };

      }
    )

    .reverse();
}


/**
 * باطل کردن کد اتصال
 */
function revokeVIPConnectionCode(
  code
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(15000);

  try {

    const cleanCode =
      normalizeVIPConnectionCode_(
        code
      );


    if (!cleanCode) {

      throw new Error(
        'کد اتصال وارد نشده است.'
      );
    }


    const sheet =
      getVIPConnectionSheet_();


    const row =
      findVIPConnectionCodeRow_(
        sheet,
        cleanCode
      );


    if (
      row === -1
    ) {

      throw new Error(
        'کد اتصال پیدا نشد.'
      );
    }


    const status =
      String(
        sheet
          .getRange(
            row,
            3
          )
          .getValue() || ''
      ).trim();


    if (
      status !==
      'صادرشده'
    ) {

      throw new Error(
        'این کد دیگر قابل ابطال نیست.'
      );
    }


    sheet
      .getRange(
        row,
        3
      )
      .setValue(
        'باطل‌شده'
      );


    return {

      success: true,

      code:
        cleanCode,

      status:
        'باطل‌شده'
    };


  } finally {

    lock.releaseLock();
  }
}


/**
 * اتصال Telegram ID به مشتری
 * با استفاده از کد یک‌بارمصرف
 */
function connectVIPTelegramAccount_(
  telegramId,
  connectionCode
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(15000);

  try {

    const cleanTelegramId =
      normalizeVIPTelegramId_(
        telegramId
      );


    const cleanCode =
      normalizeVIPConnectionCode_(
        connectionCode
      );


    if (!cleanTelegramId) {

      throw new Error(
        'شناسه تلگرام معتبر نیست.'
      );
    }


    if (!cleanCode) {

      throw new Error(
        'کد اتصال وارد نشده است.'
      );
    }


    const ss =
      SpreadsheetApp.openById(
        VIP_SHEET_ID
      );


    const customerSheet =
      ss.getSheetByName(
        CUSTOMERS_SHEET
      );


    if (!customerSheet) {

      throw new Error(
        'شیت مشتریان پیدا نشد.'
      );
    }


    const connectionSheet =
      getVIPConnectionSheet_();


    const codeRow =
      findVIPConnectionCodeRow_(
        connectionSheet,
        cleanCode
      );


    if (
      codeRow === -1
    ) {

      throw new Error(
        'کد اتصال نامعتبر است.'
      );
    }


    const codeData =
      connectionSheet
        .getRange(
          codeRow,
          1,
          1,
          8
        )
        .getValues()[0];


    const codeCustomerId =
      normalizeVIPCustomerId_(
        codeData[1]
      );


    const codeStatus =
      String(
        codeData[2] || ''
      ).trim();


    if (
      codeStatus !==
      'صادرشده'
    ) {

      throw new Error(
        'این کد اتصال قبلاً استفاده یا باطل شده است.'
      );
    }


    if (!codeCustomerId) {

      throw new Error(
        'کد اتصال به مشتری متصل نیست.'
      );
    }


    /*
     * پیدا کردن مشتری
     *
     * مهم:
     * findCustomerRow_ در Code.gs فقط
     * customerId را دریافت می‌کند.
     */
    const customerResult =
      findCustomerRow_(
        codeCustomerId
      );


    const customerRow =
      getVIPCustomerRowNumber_(
        customerResult
      );


    if (
      customerRow === -1
    ) {

      throw new Error(
        'مشتری مربوط به این کد پیدا نشد.'
      );
    }


    const customerData =
      customerSheet
        .getRange(
          customerRow,
          1,
          1,
          8
        )
        .getValues()[0];


    const customerStatus =
      customerData[7];


    if (
      typeof isVIPActive_ ===
        'function' &&
      !isVIPActive_(
        customerStatus
      )
    ) {

      throw new Error(
        'عضویت VIP این مشتری فعال نیست.'
      );
    }


    const existingTelegramId =
      normalizeVIPTelegramId_(
        customerData[4]
      );


    /*
     * اگر همین حساب قبلاً
     * به همین مشتری وصل شده باشد
     */
    if (
      existingTelegramId &&
      existingTelegramId ===
        cleanTelegramId
    ) {

      /*
       * کد را هم مصرف‌شده می‌کنیم
       * تا دوباره قابل استفاده نباشد.
       */
      const now =
        new Date();


      const displayDate =
        typeof formatPersianDateTime ===
          'function'

          ? formatPersianDateTime(
              now
            )

          : Utilities.formatDate(
              now,
              Session.getScriptTimeZone(),
              'yyyy/MM/dd HH:mm:ss'
            );


      connectionSheet
        .getRange(
          codeRow,
          3
        )
        .setValue(
          'استفاده‌شده'
        );


      connectionSheet
        .getRange(
          codeRow,
          5
        )
        .setValue(
          displayDate
        );


      connectionSheet
        .getRange(
          codeRow,
          6
        )
        .setValue(
          cleanTelegramId
        );


      connectionSheet
        .getRange(
          codeRow,
          8
        )
        .setValue(
          now
        );


      return {

        success: true,

        customerId:
          codeCustomerId,

        alreadyConnected:
          true
      };
    }


    /*
     * بررسی اینکه Telegram ID
     * قبلاً متعلق به VIP دیگری نباشد.
     */
    const lastCustomerRow =
      customerSheet.getLastRow();


    if (
      lastCustomerRow >= 2
    ) {

      const customerValues =
        customerSheet
          .getRange(
            2,
            1,
            lastCustomerRow - 1,
            8
          )
          .getValues();


      for (
        let i = 0;
        i < customerValues.length;
        i++
      ) {

        const rowTelegramId =
          normalizeVIPTelegramId_(
            customerValues[i][4]
          );


        const rowCustomerId =
          normalizeVIPCustomerId_(
            customerValues[i][0]
          );


        if (
          rowTelegramId ===
            cleanTelegramId &&
          rowCustomerId !==
            codeCustomerId
        ) {

          throw new Error(
            'این حساب تلگرام قبلاً به یک عضویت VIP دیگر متصل شده است.'
          );
        }
      }
    }


    /*
     * جلوگیری از جایگزینی Telegram ID قبلی
     */
    if (
      existingTelegramId &&
      existingTelegramId !==
        cleanTelegramId
    ) {

      throw new Error(
        'این عضویت VIP قبلاً به یک حساب تلگرام دیگر متصل شده است.'
      );
    }


    /*
     * ثبت Telegram ID در ستون E
     */
    customerSheet
      .getRange(
        customerRow,
        5
      )
      .setValue(
        cleanTelegramId
      );


    const now =
      new Date();


    const displayDate =
      typeof formatPersianDateTime ===
        'function'

        ? formatPersianDateTime(
            now
          )

        : Utilities.formatDate(
            now,
            Session.getScriptTimeZone(),
            'yyyy/MM/dd HH:mm:ss'
          );


    /*
     * مصرف کد
     */
    connectionSheet
      .getRange(
        codeRow,
        3
      )
      .setValue(
        'استفاده‌شده'
      );


    connectionSheet
      .getRange(
        codeRow,
        5
      )
      .setValue(
        displayDate
      );


    connectionSheet
      .getRange(
        codeRow,
        6
      )
      .setValue(
        cleanTelegramId
      );


    connectionSheet
      .getRange(
        codeRow,
        8
      )
      .setValue(
        now
      );


    return {

      success: true,

      customerId:
        codeCustomerId,

      alreadyConnected:
        false
    };


  } finally {

    lock.releaseLock();
  }
}
