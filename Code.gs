const HEADER_ROW = ['Current Time', 'Recorded Usage'];
const MAX_USAGE_INCREASE_KWH = 1000;
const RECENT_RECORD_COUNT = 5;
const CHART_PERIOD_DAYS = 1;
const CHART_PERIOD_COUNT = 60;
const MONTH_START_DAY = 8;
const DAY_MS = 24 * 60 * 60 * 1000;
const SUCCESS_MESSAGE = '기록을 저장했습니다.';
const UPDATE_SUCCESS_MESSAGE = '마지막 기록을 수정했습니다.';
const NO_RECORD_MESSAGE = '수정할 기록이 없습니다.';
const RANGE_ERROR_MESSAGE = '입력값이 이전 기록보다 작거나 허용 범위를 초과했습니다.';
const NUMBER_ERROR_MESSAGE = '올바른 전기 사용량 숫자를 입력해 주세요.';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('전기 사용량 기록')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function recordUsage(rawUsage) {
  const usage = parseUsage_(rawUsage);

  if (usage === null) {
    return {
      success: false,
      message: NUMBER_ERROR_MESSAGE,
    };
  }

  const sheet = getRecordingSheet_();
  ensureHeaderRow_(sheet);

  const previousUsage = getLatestRecordedUsage_(sheet);

  if (!isUsageInAllowedRange_(usage, previousUsage)) {
    return {
      success: false,
      message: RANGE_ERROR_MESSAGE,
    };
  }

  sheet.appendRow([getCurrentTimestamp_(), usage]);

  return {
    success: true,
    message: SUCCESS_MESSAGE,
    usage: usage,
  };
}

function getLatestUsage() {
  const sheet = getRecordingSheet_();
  ensureHeaderRow_(sheet);
  const records = getAllRecords_(sheet);

  return {
    success: true,
    usage: getLatestUsageFromRecords_(records),
    records: getRecentRecords_(records, RECENT_RECORD_COUNT),
    stats: getUsageStats_(records),
  };
}

function updateLatestUsage(rawUsage) {
  const usage = parseUsage_(rawUsage);

  if (usage === null) {
    return {
      success: false,
      message: NUMBER_ERROR_MESSAGE,
    };
  }

  const sheet = getRecordingSheet_();
  ensureHeaderRow_(sheet);
  const recordedUsageValues = getRecordedUsageValues_(sheet);
  const latestUsageIndex = getLatestRecordedUsageIndex_(recordedUsageValues);

  if (latestUsageIndex === -1) {
    return {
      success: false,
      message: NO_RECORD_MESSAGE,
    };
  }

  const previousUsage = getLatestRecordedUsageInValues_(recordedUsageValues, latestUsageIndex - 1);

  if (!isUsageInAllowedRange_(usage, previousUsage)) {
    return {
      success: false,
      message: RANGE_ERROR_MESSAGE,
    };
  }

  const latestUsageRow = latestUsageIndex + 2;
  sheet.getRange(latestUsageRow, 1, 1, HEADER_ROW.length)
    .setValues([[getCurrentTimestamp_(), usage]]);

  return {
    success: true,
    message: UPDATE_SUCCESS_MESSAGE,
    usage: usage,
  };
}

function parseUsage_(rawUsage) {
  if (rawUsage === null || rawUsage === undefined) {
    return null;
  }

  const trimmedUsage = String(rawUsage).trim();

  if (trimmedUsage === '') {
    return null;
  }

  const usage = Number(trimmedUsage);

  if (!Number.isFinite(usage)) {
    return null;
  }

  return usage;
}

function getRecordingSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function ensureHeaderRow_(sheet) {
  const firstRowValues = sheet.getRange(1, 1, 1, HEADER_ROW.length).getValues()[0];
  const hasExpectedHeader = HEADER_ROW.every(function (heading, index) {
    return firstRowValues[index] === heading;
  });

  if (!hasExpectedHeader) {
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
  }
}

function getRecordedUsageValues_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
}

function getLatestRecordedUsage_(sheet) {
  return getLatestRecordedUsageInValues_(getRecordedUsageValues_(sheet));
}

function getLatestRecordedUsageInValues_(recordedUsageValues, startIndex) {
  const firstIndex = startIndex === undefined ? recordedUsageValues.length - 1 : startIndex;

  for (let index = firstIndex; index >= 0; index -= 1) {
    const usage = Number(recordedUsageValues[index]);

    if (Number.isFinite(usage)) {
      return usage;
    }
  }

  return null;
}

function getLatestRecordedUsageIndex_(recordedUsageValues) {
  for (let index = recordedUsageValues.length - 1; index >= 0; index -= 1) {
    const usage = Number(recordedUsageValues[index]);

    if (Number.isFinite(usage)) {
      return index;
    }
  }

  return -1;
}

function getRecentRecords_(records, limit) {
  const startIndex = Math.max(records.length - limit, 0);

  return records.slice(startIndex).reverse().map(function (record, reverseIndex) {
    const recordIndex = records.length - 1 - reverseIndex;
    const previousRecord = recordIndex > 0 ? records[recordIndex - 1] : null;

    return {
      timestamp: formatTimestamp_(record.timestamp),
      usage: record.usage,
      dailyUsage: getDailyUsageSincePreviousRecord_(record, previousRecord),
    };
  });
}

function getDailyUsageSincePreviousRecord_(record, previousRecord) {
  if (!previousRecord) {
    return null;
  }

  const elapsedMs = record.timestamp.getTime() - previousRecord.timestamp.getTime();

  if (elapsedMs <= 0) {
    return null;
  }

  return (record.usage - previousRecord.usage) / (elapsedMs / DAY_MS);
}

function getAllRecords_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const rowValues = sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length).getValues();
  const records = [];

  for (let index = 0; index < rowValues.length; index += 1) {
    const timestamp = parseTimestamp_(rowValues[index][0]);
    const usage = Number(rowValues[index][1]);

    if (timestamp && Number.isFinite(usage)) {
      records.push({
        timestamp: timestamp,
        usage: usage,
      });
    }
  }

  records.sort(function (left, right) {
    return left.timestamp.getTime() - right.timestamp.getTime();
  });

  return records;
}

function getLatestUsageFromRecords_(records) {
  if (!records.length) {
    return null;
  }

  return records[records.length - 1].usage;
}

function getUsageStats_(records) {
  const now = new Date();
  const currentMonthStart = getCurrentUsageMonthStart_(now);
  const nextMonthStart = new Date(
    currentMonthStart.getFullYear(),
    currentMonthStart.getMonth() + 1,
    MONTH_START_DAY
  );
  const previousMonthStart = new Date(
    currentMonthStart.getFullYear(),
    currentMonthStart.getMonth() - 1,
    MONTH_START_DAY
  );

  return {
    monthStartDay: MONTH_START_DAY,
    currentMonthEstimate: getCurrentMonthEstimate_(records, currentMonthStart, nextMonthStart),
    previousMonthUsage: getPreviousMonthUsage_(records, previousMonthStart, currentMonthStart),
    recentUsageChart: getRecentUsageChart_(records, CHART_PERIOD_DAYS, CHART_PERIOD_COUNT),
  };
}

function getCurrentUsageMonthStart_(date) {
  const monthOffset = date.getDate() >= MONTH_START_DAY ? 0 : -1;
  return new Date(date.getFullYear(), date.getMonth() + monthOffset, MONTH_START_DAY);
}

function getCurrentMonthEstimate_(records, monthStart, nextMonthStart) {
  const lastRecord = getLastRecordInRange_(records, monthStart, nextMonthStart);
  let startUsage = interpolateUsageAt_(records, monthStart);
  let startTime = monthStart;

  if (startUsage === null) {
    const firstMonthRecord = getFirstRecordInRange_(records, monthStart, nextMonthStart);

    if (firstMonthRecord) {
      startUsage = firstMonthRecord.usage;
      startTime = firstMonthRecord.timestamp;
    }
  }

  if (startUsage === null || !lastRecord) {
    return {
      monthLabel: getMonthLabel_(monthStart),
      value: null,
    };
  }

  const elapsedMs = lastRecord.timestamp.getTime() - startTime.getTime();

  if (elapsedMs <= 0) {
    return {
      monthLabel: getMonthLabel_(monthStart),
      value: null,
    };
  }

  const monthMs = nextMonthStart.getTime() - monthStart.getTime();
  const estimatedUsage = (lastRecord.usage - startUsage) / elapsedMs * monthMs;

  return {
    monthLabel: getMonthLabel_(monthStart),
    value: estimatedUsage,
  };
}

function getPreviousMonthUsage_(records, monthStart, nextMonthStart) {
  const endUsage = interpolateUsageAt_(records, nextMonthStart);
  let startUsage = interpolateUsageAt_(records, monthStart);
  let startTime = monthStart;
  let shouldEstimateFullMonth = false;

  if (startUsage === null) {
    const firstMonthRecord = getFirstRecordInRange_(records, monthStart, nextMonthStart);

    if (firstMonthRecord) {
      startUsage = firstMonthRecord.usage;
      startTime = firstMonthRecord.timestamp;
      shouldEstimateFullMonth = true;
    }
  }

  if (startUsage === null || endUsage === null) {
    return {
      monthLabel: getMonthLabel_(monthStart),
      value: null,
    };
  }

  if (shouldEstimateFullMonth) {
    const elapsedMs = nextMonthStart.getTime() - startTime.getTime();
    const monthMs = nextMonthStart.getTime() - monthStart.getTime();

    if (elapsedMs <= 0) {
      return {
        monthLabel: getMonthLabel_(monthStart),
        value: null,
      };
    }

    return {
      monthLabel: getMonthLabel_(monthStart),
      value: (endUsage - startUsage) / elapsedMs * monthMs,
    };
  }

  return {
    monthLabel: getMonthLabel_(monthStart),
    value: endUsage - startUsage,
  };
}

function getFirstRecordInRange_(records, startTime, endTime) {
  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];

    if (record.timestamp >= startTime && record.timestamp < endTime) {
      return record;
    }
  }

  return null;
}

function getLastRecordInRange_(records, startTime, endTime) {
  for (let index = records.length - 1; index >= 0; index -= 1) {
    const record = records[index];

    if (record.timestamp >= startTime && record.timestamp < endTime) {
      return record;
    }
  }

  return null;
}

function getRecentUsageChart_(records, periodDays, periodCount) {
  if (!records.length) {
    return {
      periodDays: periodDays,
      periodCount: periodCount,
      points: [],
    };
  }

  const latestRecord = records[records.length - 1];
  const latestTime = latestRecord.timestamp;
  const periodMs = periodDays * DAY_MS;
  const points = [];

  for (let index = periodCount - 1; index >= 0; index -= 1) {
    const startTime = new Date(latestTime.getTime() - (index + 1) * periodMs);
    const endTime = new Date(latestTime.getTime() - index * periodMs);
    const startUsage = interpolateUsageAt_(records, startTime);
    const endUsage = interpolateUsageAt_(records, endTime);

    if (startUsage === null || endUsage === null) {
      continue;
    }

    points.push({
      label: getCompactChartDateLabel_(endTime),
      periodLabel: getCompactChartDateLabel_(startTime) + '~' + getCompactChartDateLabel_(endTime),
      startTimestamp: formatTimestamp_(startTime),
      endTimestamp: formatTimestamp_(endTime),
      dailyUsage: (endUsage - startUsage) / periodDays,
    });
  }

  return {
    periodDays: periodDays,
    periodCount: periodCount,
    points: points,
  };
}

function interpolateUsageAt_(records, targetTime) {
  let beforeRecord = null;
  let afterRecord = null;
  const targetMs = targetTime.getTime();

  for (let index = 0; index < records.length; index += 1) {
    const record = records[index];
    const recordMs = record.timestamp.getTime();

    if (recordMs === targetMs) {
      return record.usage;
    }

    if (recordMs < targetMs) {
      beforeRecord = record;
      continue;
    }

    afterRecord = record;
    break;
  }

  if (!beforeRecord || !afterRecord) {
    return null;
  }

  const beforeMs = beforeRecord.timestamp.getTime();
  const afterMs = afterRecord.timestamp.getTime();

  if (afterMs === beforeMs) {
    return null;
  }

  return beforeRecord.usage +
    (afterRecord.usage - beforeRecord.usage) *
    ((targetMs - beforeMs) / (afterMs - beforeMs));
}

function parseTimestamp_(timestamp) {
  if (timestamp instanceof Date) {
    return timestamp;
  }

  const text = String(timestamp || '').trim();
  const matches = text.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})$/);

  if (!matches) {
    return null;
  }

  return new Date(
    Number(matches[1]),
    Number(matches[2]) - 1,
    Number(matches[3]),
    Number(matches[4]),
    Number(matches[5]),
    Number(matches[6])
  );
}

function formatTimestamp_(timestamp) {
  if (timestamp instanceof Date) {
    return Utilities.formatDate(
      timestamp,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
  }

  return String(timestamp || '');
}

function getMonthLabel_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy년 M월');
}

function getCompactChartDateLabel_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMdd');
}

function isUsageInAllowedRange_(usage, previousUsage) {
  if (previousUsage === null) {
    return true;
  }

  return usage >= previousUsage && usage <= previousUsage + MAX_USAGE_INCREASE_KWH;
}

function getCurrentTimestamp_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}
