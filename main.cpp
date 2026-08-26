#include <QApplication>
#include <QCoreApplication>
#include <QDate>
#include <QDateTime>
#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QFormLayout>
#include <QHBoxLayout>
#include <QLineEdit>
#include <QMessageBox>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QRegularExpression>
#include <QString>
#include <QStringList>
#include <QTextStream>
#include <QTime>
#include <QVariant>
#include <QVBoxLayout>
#include <QVector>
#include <QWidget>
#include <limits>

// QXlsx reads/writes .xlsx files directly.
// The program does not use Microsoft Excel COM and does not start Excel.
#include "xlsxcellrange.h"
#include "xlsxdocument.h"

using namespace QXlsx;

namespace {

// QXlsx uses Excel's 1-based row/column numbers:
// A=1, B=2, ..., BI=61.
constexpr int kFirstAverageCol = 2;  // B column: first component data column.
constexpr int kLastAverageCol = 61;  // BI column: last component data column.
constexpr int kRequiredSampleCount = 30;
constexpr int kSingleRowSampleCount = 1;

// Parsed value of one Excel cell.
// text is used for labels such as "Data:", "Pin:", and "Conditions:".
// numeric/number are used for average calculation.
struct CellValue {
    QString text;
    bool numeric = false;
    double number = 0.0;
};

// One row in the output workbook.
// averages/maxValues store one result per source column from B to BI.
struct ConditionAverage {
    QString sheetName;
    QString testItem;
    QString testStartTime;
    QString testEndTime;
    QString condition;
    QString marker;
    int markerRow = 0;
    int averageSampleStart = 0;
    int averageSampleEnd = 0;
    int maxSampleStart = 0;
    int maxSampleEnd = 0;
    QVector<double> averages;
    QVector<double> maxValues;
};

struct FanResult {
    QString sheetName;
    QString testItem;
    QString testStartTime;
    QString testEndTime;
    QString condition;
    QString marker;
    int markerRow = 0;
    int lastDataRow = 0;
    int maxSampleStart = 0;
    int maxSampleEnd = 0;
    QVector<double> lastValues;
    QVector<double> maxValues;
};

struct SheetData {
    QString name;
    QVector<QVector<CellValue>> rows;
    int lastRow = 0;
};

QString cellDisplayText(const QVariant &value)
{
    const QDateTime dateTime = value.toDateTime();
    if (dateTime.isValid())
        return dateTime.toString("yyyy/M/d HH:mm");

    const QDate date = value.toDate();
    if (date.isValid())
        return date.toString("yyyy/M/d");

    const QTime time = value.toTime();
    if (time.isValid())
        return time.toString("HH:mm");

    return value.toString().trimmed();
}

CellValue parseCell(const QVariant &value)
{
    CellValue cell;
    // Convert every cell to trimmed text for stable label matching.
    // Example: " Pin: " becomes "Pin:".
    cell.text = cellDisplayText(value).trimmed();

    // Numeric cells may be stored as int/double QVariant.
    // Numeric strings are also accepted because some Excel files store numbers as text.
    bool ok = false;
    double number = value.toDouble(&ok);
    if (!ok && !cell.text.isEmpty())
        number = cell.text.toDouble(&ok);

    cell.numeric = ok;
    cell.number = ok ? number : 0.0;
    return cell;
}

bool isDataRow(const QVector<CellValue> &row)
{
    const QString colA = row.value(1).text;
    const CellValue colB = row.value(2);

    // Thermal reports can mark the first data row with A="Data:".
    // The following rows in the same block may leave A blank, while B:BI still
    // contain numeric measurement values.
    return colA == "Data:" || (colA.isEmpty() && colB.numeric);
}

bool isConditionLabel(const QString &text)
{
    return text == "Conditions:" || text == "Condition:";
}

bool isTestItemRow(const QVector<CellValue> &row)
{
    return row.value(1).text == "Test Item Name:";
}

bool isTestTimeRow(const QVector<CellValue> &row)
{
    return row.value(1).text == "Test Time:";
}

QString firstTextInRow(const QVector<CellValue> &row, int firstCol, int lastCol)
{
    for (int c = firstCol; c <= lastCol; ++c) {
        const QString text = row.value(c).text;
        if (!text.isEmpty())
            return text;
    }
    return {};
}

QString findBaseCondition(const QVector<QVector<CellValue>> &rows, int row)
{
    // A B-column marker row marks the end of one test-condition block.
    // Search upward to find the original condition text from column B.
    // Accepted source rows:
    //   A="Conditions:" or A="Condition:", B="<condition text>"
    for (int r = row; r >= 1; --r) {
        const CellValue colB = rows.value(r).value(2);
        if (isConditionLabel(rows.value(r).value(1).text) && !colB.text.isEmpty())
            return colB.text;
        if (r != row && isTestItemRow(rows.value(r)))
            break;
    }
    return {};
}

QString extractPercentText(const QString &text)
{
    // Example: "70% OCP Trip Loading Value:" -> "70%".
    static const QRegularExpression percentPattern(R"((\d+(?:\.\d+)?)\s*%)");
    const QRegularExpressionMatch match = percentPattern.match(text);
    if (!match.hasMatch())
        return {};

    return match.captured(1) + "%";
}

QString findTestItem(const QVector<QVector<CellValue>> &rows, int row);

QString findOcpNotWorkingCondition(const QVector<QVector<CellValue>> &rows, int markerRow)
{
    // OCP Not Working is tied to the nearest OCP trip-loading row above it.
    // Single Channel reports put "OCP Load Channel:" in column B, while
    // ALL Channel reports keep the same percentage in column A but leave
    // column B as the numeric loading value.
    for (int r = markerRow - 1; r >= 1; --r) {
        if (isTestItemRow(rows.value(r)))
            break;

        const QString percent = extractPercentText(rows.value(r).value(1).text);
        if (percent.isEmpty())
            continue;

        QString baseCondition = findBaseCondition(rows, r);
        if (baseCondition.isEmpty())
            baseCondition = findTestItem(rows, markerRow);
        if (percent.isEmpty())
            return baseCondition;
        return baseCondition + " " + percent;
    }

    return findBaseCondition(rows, markerRow);
}

QString findCondition(const QVector<QVector<CellValue>> &rows, int markerRow, const QString &marker)
{
    if (marker == "OCP Not Working")
        return findOcpNotWorkingCondition(rows, markerRow);

    return findBaseCondition(rows, markerRow);
}

QString findTestItem(const QVector<QVector<CellValue>> &rows, int row)
{
    for (int r = row; r >= 1; --r) {
        if (isTestItemRow(rows.value(r)))
            return rows.value(r).value(2).text;
    }
    return {};
}

QString findTestStartTime(const QVector<QVector<CellValue>> &rows, int row)
{
    for (int r = row; r >= 1; --r) {
        if (isTestTimeRow(rows.value(r)))
            return rows.value(r).value(2).text;
    }
    return {};
}

QString markerEndTime(const QVector<QVector<CellValue>> &rows, int markerRow)
{
    // Marker rows normally put the end timestamp in column A.
    // If column A is a step label instead, keep it as-is because it is still
    // useful traceability in the generated report.
    return rows.value(markerRow).value(1).text;
}

QString dynamicEndTime(const QVector<QVector<CellValue>> &rows, int blockEndRow, int lastDataRow)
{
    // Thermal Dynamic Test has no Pin/OTP/OCP marker. Use the first non-data row
    // after the final Data row inside the same test block as its end record.
    for (int r = lastDataRow + 1; r <= blockEndRow && r < rows.size(); ++r) {
        if (isDataRow(rows.value(r)))
            continue;

        const QString text = firstTextInRow(rows.value(r), 1, 2);
        if (!text.isEmpty())
            return text;
    }
    return {};
}

QString findTextInBlock(const QVector<QVector<CellValue>> &rows, int blockStartRow, int blockEndRow, const QString &label)
{
    for (int r = blockStartRow; r <= blockEndRow && r < rows.size(); ++r) {
        if (rows.value(r).value(1).text == label)
            return rows.value(r).value(2).text;
    }
    return {};
}

bool isUserAbortRow(const QVector<CellValue> &row)
{
    return row.value(2).text == "User Abort Test!";
}

int testBlockStartRow(const QVector<QVector<CellValue>> &rows, int row)
{
    for (int r = row; r >= 1; --r) {
        if (isTestItemRow(rows.value(r)))
            return r;
    }
    return 1;
}

int testBlockEndRow(const QVector<QVector<CellValue>> &rows, int blockStartRow)
{
    for (int r = blockStartRow + 1; r < rows.size(); ++r) {
        if (isTestItemRow(rows.value(r)))
            return r - 1;
    }
    return rows.size() - 1;
}

bool testBlockHasUserAbort(const QVector<QVector<CellValue>> &rows, int blockStartRow, int blockEndRow)
{
    for (int r = blockStartRow; r <= blockEndRow && r < rows.size(); ++r) {
        if (isUserAbortRow(rows.value(r)))
            return true;
    }
    return false;
}

bool conditionContainsFan(const QString &condition)
{
    return condition.contains("FAN", Qt::CaseInsensitive);
}

QVector<QVector<CellValue>> readSheet(Document &book, int *lastRow)
{
    // Read the active worksheet into a 1-based 2D array.
    // Index 0 is intentionally unused so Excel row/column numbers match C++ indexes.
    const CellRange range = book.dimension();
    const int maxRow = qMax(range.lastRow(), 1);
    const int maxCol = qMax(range.lastColumn(), kLastAverageCol);

    QVector<QVector<CellValue>> rows(maxRow + 1);
    for (int r = 1; r <= maxRow; ++r) {
        rows[r].resize(maxCol + 1);
        bool hasValue = false;

        for (int c = 1; c <= maxCol; ++c) {
            rows[r][c] = parseCell(book.read(r, c));
            hasValue = hasValue || !rows[r][c].text.isEmpty();
        }

        // lastRow is used only for log display.
        if (hasValue && lastRow)
            *lastRow = r;
    }
    return rows;
}

QStringList parseCsvLine(const QString &line)
{
    QStringList fields;
    QString field;
    bool inQuotes = false;

    for (int i = 0; i < line.size(); ++i) {
        const QChar ch = line[i];
        if (ch == '"') {
            if (inQuotes && i + 1 < line.size() && line[i + 1] == '"') {
                field.append('"');
                ++i;
            } else {
                inQuotes = !inQuotes;
            }
            continue;
        }

        if (ch == ',' && !inQuotes) {
            fields.append(field);
            field.clear();
            continue;
        }

        field.append(ch);
    }

    fields.append(field);
    return fields;
}

QVector<QVector<CellValue>> readCsv(const QString &inputPath, int *lastRow, QString *error)
{
    QFile file(inputPath);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        if (error)
            *error = "無法開啟 CSV：" + inputPath;
        return {};
    }

    QVector<QVector<CellValue>> rows(1);
    QTextStream stream(&file);

    while (!stream.atEnd()) {
        QString line = stream.readLine();
        if (rows.size() == 1)
            line.remove(QChar(0xFEFF)); // Remove UTF-8 BOM if present.

        const QStringList fields = parseCsvLine(line);
        const int rowIndex = rows.size();
        rows.append(QVector<CellValue>(qMax(fields.size(), kLastAverageCol) + 1));

        bool hasValue = false;
        for (int i = 0; i < fields.size(); ++i) {
            const int col = i + 1;
            rows[rowIndex][col] = parseCell(fields[i]);
            hasValue = hasValue || !rows[rowIndex][col].text.isEmpty();
        }

        if (hasValue && lastRow)
            *lastRow = rowIndex;
    }

    return rows;
}

QVector<SheetData> readInputSheets(const QString &inputPath, int *totalLastRow, QString *error)
{
    if (totalLastRow)
        *totalLastRow = 0;

    if (inputPath.endsWith(".csv", Qt::CaseInsensitive)) {
        SheetData sheet;
        sheet.name = "CSV";
        sheet.rows = readCsv(inputPath, &sheet.lastRow, error);
        if (totalLastRow)
            *totalLastRow = sheet.lastRow;
        return sheet.rows.isEmpty() ? QVector<SheetData>{} : QVector<SheetData>{sheet};
    }

    Document inputBook(inputPath);
    if (!inputBook.load()) {
        if (error)
            *error = "無法讀取 Excel，請確認檔案格式為 .xlsx：" + inputPath;
        return {};
    }

    QVector<SheetData> sheets;
    const QStringList sheetNames = inputBook.sheetNames();
    for (const QString &sheetName : sheetNames) {
        if (sheetName.compare("Report Template", Qt::CaseInsensitive) == 0)
            continue;
        if (!inputBook.selectSheet(sheetName))
            continue;

        SheetData sheet;
        sheet.name = sheetName;
        sheet.rows = readSheet(inputBook, &sheet.lastRow);
        if (totalLastRow)
            *totalLastRow += sheet.lastRow;
        sheets.append(sheet);
    }

    if (sheets.isEmpty() && error)
        *error = "沒有可處理的資料工作表。";

    return sheets;
}

int averageSampleCountForMarker(const QString &marker)
{
    if (marker == "Pin:")
        return kRequiredSampleCount;
    if (marker == "Release Short State")
        return kRequiredSampleCount;
    if (marker == "OTP Occur")
        return kSingleRowSampleCount;
    if (marker == "OCP Not Working")
        return kRequiredSampleCount;
    return 0;
}

int maxSampleCountForMarker(const QString &marker)
{
    if (marker == "Pin:")
        return kRequiredSampleCount;
    if (marker == "Release Short State")
        return kRequiredSampleCount;
    if (marker == "OTP Occur")
        return kRequiredSampleCount;
    if (marker == "OCP Not Working")
        return kRequiredSampleCount;
    return 0;
}

bool isFailedPinMarker(const QVector<CellValue> &row)
{
    return row.value(2).text == "Pin:" && row.value(1).text == "Test FAIL!";
}

QVector<int> contiguousDataRowsAboveMarker(const QVector<QVector<CellValue>> &rows, int markerRow)
{
    QVector<int> dataRows;

    // Starting from the row immediately above the marker, walk upward while the
    // rows still belong to the same continuous data block. Stop at the first
    // non-data row after data has already been collected.
    for (int r = markerRow - 1; r >= 1; --r) {
        if (isDataRow(rows[r])) {
            dataRows.prepend(r);
            continue;
        }
        if (!dataRows.isEmpty())
            break;
    }
    return dataRows;
}

ConditionAverage averageBlock(const QVector<QVector<CellValue>> &rows,
                              const QVector<int> &averageRows,
                              const QVector<int> &maxRows,
                              int markerRow,
                              const QString &marker,
                              const QString &sheetName,
                              const QString &testEndTime = {})
{
    ConditionAverage result;
    result.sheetName = sheetName;
    result.testItem = findTestItem(rows, markerRow);
    result.testStartTime = findTestStartTime(rows, markerRow);
    result.testEndTime = testEndTime.isEmpty() ? markerEndTime(rows, markerRow) : testEndTime;
    result.condition = findCondition(rows, markerRow, marker);
    result.marker = marker;
    result.markerRow = markerRow;
    result.averageSampleStart = averageRows.first();
    result.averageSampleEnd = averageRows.last();
    result.maxSampleStart = maxRows.first();
    result.maxSampleEnd = maxRows.last();

    // Calculate average and max for each component column.
    // Columns are processed independently: B, C, ..., BI.
    for (int c = kFirstAverageCol; c <= kLastAverageCol; ++c) {
        double sum = 0.0;
        double maxValue = std::numeric_limits<double>::lowest();
        int averageCount = 0;
        int maxCount = 0;

        for (int r : averageRows) {
            const CellValue cell = rows.value(r).value(c);
            if (cell.numeric) {
                sum += cell.number;
                ++averageCount;
            }
        }

        for (int r : maxRows) {
            const CellValue cell = rows.value(r).value(c);
            if (cell.numeric) {
                maxValue = qMax(maxValue, cell.number);
                ++maxCount;
            }
        }

        // Avoid division by zero if a column has no numeric value in the sample.
        result.averages.append(averageCount == 0 ? 0.0 : sum / averageCount);
        result.maxValues.append(maxCount == 0 ? 0.0 : maxValue);
    }
    return result;
}

QVector<ConditionAverage> calculateMarkerAverages(const QVector<QVector<CellValue>> &rows, const QString &sheetName)
{
    QVector<ConditionAverage> results;

    for (int r = 2; r < rows.size(); ++r) {
        const QString marker = rows[r].value(2).text;
        const int requestedAverageSampleCount = averageSampleCountForMarker(marker);
        const int requestedMaxSampleCount = maxSampleCountForMarker(marker);

        // Real marker rule:
        // 1. Column B must equal one supported marker:
        //    "Pin:", "Release Short State", "OTP Occur", or "OCP Not Working".
        // 2. The row above it must be a data row.
        //
        // This avoids fixed row numbers and avoids matching unrelated labels
        // that may appear elsewhere in the report.
        if (requestedAverageSampleCount == 0 || requestedMaxSampleCount == 0 || isFailedPinMarker(rows[r]) || !isDataRow(rows[r - 1]))
            continue;

        const int blockStartRow = testBlockStartRow(rows, r);
        const int blockEndRow = testBlockEndRow(rows, blockStartRow);
        if (testBlockHasUserAbort(rows, blockStartRow, blockEndRow))
            continue;

        const QVector<int> dataRows = contiguousDataRowsAboveMarker(rows, r);
        const int averageSampleCount = qMin(requestedAverageSampleCount, dataRows.size());
        const int maxSampleCount = qMin(requestedMaxSampleCount, dataRows.size());
        if (averageSampleCount == 0 || maxSampleCount == 0)
            continue;

        // Core rule:
        // Pin:         -> average and max final 30 data rows.
        // Release Short State -> average and max final 30 data rows.
        // OTP Occur    -> Average uses the row immediately above the marker;
        //                 Max uses the final 30 data rows.
        // OCP Not Working -> average and max final 30 data rows.
        // If the block has fewer rows than requested, use all available rows.
        const QVector<int> averageRows = dataRows.mid(dataRows.size() - averageSampleCount, averageSampleCount);
        const QVector<int> maxRows = dataRows.mid(dataRows.size() - maxSampleCount, maxSampleCount);
        results.append(averageBlock(rows, averageRows, maxRows, r, marker, sheetName));
    }
    return results;
}

QVector<ConditionAverage> calculateDynamicAverages(const QVector<QVector<CellValue>> &rows, const QString &sheetName)
{
    QVector<ConditionAverage> results;

    for (int r = 1; r < rows.size(); ++r) {
        if (!isTestItemRow(rows.value(r)) || rows.value(r).value(2).text != "Thermal Dynamic Test")
            continue;

        const int blockEndRow = testBlockEndRow(rows, r);
        if (testBlockHasUserAbort(rows, r, blockEndRow))
            continue;

        QVector<int> dataRows;
        for (int row = r + 1; row <= blockEndRow; ++row) {
            if (isDataRow(rows.value(row)))
                dataRows.append(row);
        }

        const int sampleCount = qMin(kRequiredSampleCount, dataRows.size());
        if (sampleCount == 0)
            continue;

        const QVector<int> sampleRows = dataRows.mid(dataRows.size() - sampleCount, sampleCount);
        const int lastDataRow = sampleRows.last();
        const QString endTime = dynamicEndTime(rows, blockEndRow, lastDataRow);

        ConditionAverage result = averageBlock(rows,
                                               sampleRows,
                                               sampleRows,
                                               r,
                                               "Thermal Dynamic Test",
                                               sheetName,
                                               endTime);
        result.testStartTime = findTextInBlock(rows, r, blockEndRow, "Test Time:");
        result.condition = findTextInBlock(rows, r, blockEndRow, "Condition:");
        result.markerRow = lastDataRow;
        results.append(result);
    }

    return results;
}

QVector<ConditionAverage> calculateAverages(const QVector<QVector<CellValue>> &rows, const QString &sheetName)
{
    QVector<ConditionAverage> results = calculateMarkerAverages(rows, sheetName);
    results += calculateDynamicAverages(rows, sheetName);
    return results;
}

FanResult fanBlock(const QVector<QVector<CellValue>> &rows,
                   const QVector<int> &lastRows,
                   const QVector<int> &maxRows,
                   int markerRow,
                   const QString &sheetName)
{
    FanResult result;
    result.sheetName = sheetName;
    result.testItem = findTestItem(rows, markerRow);
    result.testStartTime = findTestStartTime(rows, markerRow);
    result.testEndTime = markerEndTime(rows, markerRow);
    result.condition = findBaseCondition(rows, markerRow);
    result.marker = "Vout Abnormal";
    result.markerRow = markerRow;
    result.lastDataRow = lastRows.last();
    result.maxSampleStart = maxRows.first();
    result.maxSampleEnd = maxRows.last();

    for (int c = kFirstAverageCol; c <= kLastAverageCol; ++c) {
        const CellValue lastCell = rows.value(result.lastDataRow).value(c);
        result.lastValues.append(lastCell.numeric ? lastCell.number : 0.0);

        double maxValue = std::numeric_limits<double>::lowest();
        int maxCount = 0;
        for (int r : maxRows) {
            const CellValue cell = rows.value(r).value(c);
            if (cell.numeric) {
                maxValue = qMax(maxValue, cell.number);
                ++maxCount;
            }
        }
        result.maxValues.append(maxCount == 0 ? 0.0 : maxValue);
    }
    return result;
}

QVector<FanResult> calculateFanResults(const QVector<QVector<CellValue>> &rows, const QString &sheetName)
{
    QVector<FanResult> results;

    for (int r = 2; r < rows.size(); ++r) {
        if (rows[r].value(2).text != "Vout Abnormal" || !isDataRow(rows[r - 1]))
            continue;

        const int blockStartRow = testBlockStartRow(rows, r);
        const int blockEndRow = testBlockEndRow(rows, blockStartRow);
        if (testBlockHasUserAbort(rows, blockStartRow, blockEndRow))
            continue;

        const QString condition = findBaseCondition(rows, r);
        if (!conditionContainsFan(condition))
            continue;

        const QVector<int> dataRows = contiguousDataRowsAboveMarker(rows, r);
        const int lastSampleCount = qMin(kSingleRowSampleCount, dataRows.size());
        const int maxSampleCount = qMin(kRequiredSampleCount, dataRows.size());
        if (lastSampleCount == 0 || maxSampleCount == 0)
            continue;

        const QVector<int> lastRows = dataRows.mid(dataRows.size() - lastSampleCount, lastSampleCount);
        const QVector<int> maxRows = dataRows.mid(dataRows.size() - maxSampleCount, maxSampleCount);
        results.append(fanBlock(rows, lastRows, maxRows, r, sheetName));
    }
    return results;
}

void writeTemplateLabels(Document &outBook, const QString &sheetName)
{
    outBook.selectSheet(sheetName);

    // Report Template layout from the user's sample workbook.
    // Row 1..4 describe the test block; row 5..64 map to source B..BI as 1..60.
    outBook.write(1, 1, "Test Item:");
    outBook.write(2, 1, "Test Start Time:");
    outBook.write(3, 1, "Test End Time:");
    outBook.write(4, 1, "Condition:");
    for (int i = 1; i <= kLastAverageCol - kFirstAverageCol + 1; ++i)
        outBook.write(i + 4, 1, i);

    outBook.setColumnWidth(1, 18);
}

void writeTemplateResultSheet(Document &outBook,
                              const QString &sheetName,
                              const QVector<ConditionAverage> &results,
                              bool writeMaxValues)
{
    writeTemplateLabels(outBook, sheetName);

    int outputCol = 2;
    for (int i = 0; i < results.size(); ++i) {
        if (!writeMaxValues && results[i].marker == "OTP Occur")
            continue;

        outBook.write(1, outputCol, results[i].testItem);
        outBook.write(2, outputCol, results[i].testStartTime);
        outBook.write(3, outputCol, results[i].testEndTime);
        outBook.write(4, outputCol, results[i].condition);
        outBook.setColumnWidth(outputCol, 24);

        const QVector<double> &values = writeMaxValues ? results[i].maxValues : results[i].averages;
        for (int c = 0; c < values.size(); ++c)
            outBook.write(c + 5, outputCol, values[c]);

        ++outputCol;
    }
}

void writeFanSheet(Document &outBook, const QVector<FanResult> &results)
{
    writeTemplateLabels(outBook, "FAN");

    for (int i = 0; i < results.size(); ++i) {
        const int outputCol = i + 2;
        outBook.write(1, outputCol, results[i].testItem);
        outBook.write(2, outputCol, results[i].testStartTime);
        outBook.write(3, outputCol, results[i].testEndTime);
        outBook.write(4, outputCol, results[i].condition);
        outBook.setColumnWidth(outputCol, 24);

        // FAN lock output keeps the current rule: use the final data row before
        // Vout Abnormal, but present it with the same template layout as
        // Average/Max.
        for (int c = 0; c < results[i].lastValues.size(); ++c)
            outBook.write(c + 5, outputCol, results[i].lastValues[c]);
    }
}

bool writeResults(const QString &outputPath,
                  const QVector<ConditionAverage> &results,
                  const QVector<FanResult> &fanResults)
{
    Document outBook;
    outBook.addSheet("Average");
    outBook.addSheet("Max");

    writeTemplateResultSheet(outBook, "Average", results, false);
    writeTemplateResultSheet(outBook, "Max", results, true);
    if (!fanResults.isEmpty()) {
        outBook.addSheet("FAN");
        writeFanSheet(outBook, fanResults);
    }

    return outBook.saveAs(outputPath);
}

bool processWorkbook(const QString &inputPath,
                     const QString &outputPath,
                     QString *log,
                     QString *error)
{
    // Helper lambdas keep the workflow readable.
    // log/error are pointers so both GUI mode and command-line mode can reuse
    // the same processing function.
    auto appendLog = [log](const QString &line) {
        if (log)
            log->append(line + "\n");
    };
    auto fail = [error](const QString &message) {
        if (error)
            *error = message;
        return false;
    };

    if (!QFile::exists(inputPath))
        return fail("找不到輸入檔：" + inputPath);

    const bool isXlsx = inputPath.endsWith(".xlsx", Qt::CaseInsensitive);
    const bool isCsv = inputPath.endsWith(".csv", Qt::CaseInsensitive);
    if (!isXlsx && !isCsv)
        return fail("目前只支援 .xlsx 或 .csv，請先將來源另存成支援格式。");

    // If the output file is open in Excel, Windows usually prevents removal.
    // Check this early so the user gets a clear message.
    if (QFile::exists(outputPath) && !QFile::remove(outputPath))
        return fail("輸出檔案無法覆蓋，請確認檔案沒有被 Excel 開啟：" + outputPath);

    int totalLastRow = 0;
    QString readError;
    const QVector<SheetData> sheets = readInputSheets(inputPath, &totalLastRow, &readError);
    if (!readError.isEmpty())
        return fail(readError);

    QVector<ConditionAverage> results;
    QVector<FanResult> fanResults;
    for (const SheetData &sheet : sheets) {
        results += calculateAverages(sheet.rows, sheet.name);
        fanResults += calculateFanResults(sheet.rows, sheet.name);
    }

    appendLog("Input: " + inputPath);
    appendLog("Data sheets read: " + QString::number(sheets.size()));
    appendLog("Used rows read total: " + QString::number(totalLastRow));
    appendLog("Detected conditions: " + QString::number(results.size()));
    for (int i = 0; i < results.size(); ++i) {
        const ConditionAverage &result = results[i];
        appendLog(QString("  #%1 Sheet=%2 Marker=%3 MarkerRow=%4 AvgRows=%5:%6 MaxRows=%7:%8 Condition=%9")
                      .arg(i + 1)
                      .arg(result.sheetName)
                      .arg(result.marker)
                      .arg(result.markerRow)
                      .arg(result.averageSampleStart)
                      .arg(result.averageSampleEnd)
                      .arg(result.maxSampleStart)
                      .arg(result.maxSampleEnd)
                      .arg(result.condition));
    }
    appendLog("Detected FAN Vout Abnormal: " + QString::number(fanResults.size()));
    for (int i = 0; i < fanResults.size(); ++i) {
        const FanResult &result = fanResults[i];
        appendLog(QString("  FAN #%1 Sheet=%2 MarkerRow=%3 LastRow=%4 MaxRows=%5:%6 Condition=%7")
                      .arg(i + 1)
                      .arg(result.sheetName)
                      .arg(result.markerRow)
                      .arg(result.lastDataRow)
                      .arg(result.maxSampleStart)
                      .arg(result.maxSampleEnd)
                      .arg(result.condition));
    }
    if (results.isEmpty() && fanResults.isEmpty())
        return fail("沒有找到符合規則的資料：B 欄需為 Pin: / Release Short State / OTP Occur / OCP Not Working / Vout Abnormal，或 Test Item Name: 為 Thermal Dynamic Test。");

    if (!writeResults(outputPath, results, fanResults))
        return fail("輸出 Excel 失敗，請確認路徑可寫入：" + outputPath);

    appendLog("Output: " + outputPath);
    return true;
}

QString defaultOutputPath(const QString &inputPath)
{
    // Put the result beside the input workbook by default.
    // Example: Test.xlsx -> Test_autoReport.xlsx.
    const QFileInfo info(inputPath);
    return info.absolutePath() + "/" + info.completeBaseName() + "_autoReport.xlsx";
}

} // namespace

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);

    // Command-line mode for testing/automation:
    // ThermalTool.exe input.xlsx output.xlsx
    const QStringList args = QCoreApplication::arguments();
    if (args.size() >= 3) {
        QString log;
        QString error;

        const bool ok = processWorkbook(args[1],
                                        args[2],
                                        &log,
                                        &error);

        QTextStream(stdout) << log;
        if (!ok)
            QTextStream(stderr) << error << Qt::endl;
        return ok ? 0 : 1;
    }

    // GUI mode for normal users: choose input/output paths and press Run.
    QWidget window;
    window.setWindowTitle("Thermal Report Tool");
    window.resize(760, 420);

    auto *inputEdit = new QLineEdit("C:/Qt_Project/Thermal Report/Data/Thermal test.xlsx");
    auto *outputEdit = new QLineEdit(defaultOutputPath(inputEdit->text()));
    auto *inputButton = new QPushButton("選擇...");
    auto *outputButton = new QPushButton("選擇...");
    auto *runButton = new QPushButton("執行");
    auto *logEdit = new QPlainTextEdit;
    logEdit->setReadOnly(true);
    logEdit->setPlaceholderText("執行結果會顯示在這裡");

    auto *inputRow = new QHBoxLayout;
    inputRow->addWidget(inputEdit);
    inputRow->addWidget(inputButton);

    auto *outputRow = new QHBoxLayout;
    outputRow->addWidget(outputEdit);
    outputRow->addWidget(outputButton);

    auto *form = new QFormLayout;
    form->addRow("輸入 Excel", inputRow);
    form->addRow("輸出 Excel", outputRow);

    auto *layout = new QVBoxLayout(&window);
    layout->addLayout(form);
    layout->addWidget(runButton);
    layout->addWidget(logEdit);

    QObject::connect(inputButton, &QPushButton::clicked, [&]() {
        // Show a file picker and store the selected input workbook path.
        const QString file = QFileDialog::getOpenFileName(&window,
                                                          "選擇輸入檔",
                                                          inputEdit->text(),
                                                          "Data Files (*.xlsx *.csv)");
        if (file.isEmpty())
            return;

        inputEdit->setText(QDir::toNativeSeparators(file));

        // When input changes, auto-generate a matching output path.
        outputEdit->setText(QDir::toNativeSeparators(defaultOutputPath(file)));
    });

    QObject::connect(outputButton, &QPushButton::clicked, [&]() {
        // Let the user choose where the result workbook should be saved.
        const QString file = QFileDialog::getSaveFileName(&window,
                                                          "選擇輸出 Excel",
                                                          outputEdit->text(),
                                                          "Excel Files (*.xlsx)");
        if (!file.isEmpty())
            outputEdit->setText(QDir::toNativeSeparators(file));
    });

    QObject::connect(runButton, &QPushButton::clicked, [&]() {
        const QString inputPath = QDir::fromNativeSeparators(inputEdit->text().trimmed());
        QString outputPath = QDir::fromNativeSeparators(outputEdit->text().trimmed());

        if (inputPath.isEmpty()) {
            QMessageBox::warning(&window, "缺少輸入檔", "請先選擇輸入 Excel。");
            return;
        }

        if (outputPath.isEmpty())
            outputPath = defaultOutputPath(inputPath);
        if (!outputPath.endsWith(".xlsx", Qt::CaseInsensitive))
            outputPath += ".xlsx";
        outputEdit->setText(QDir::toNativeSeparators(outputPath));

        // Give immediate UI feedback before processing starts.
        // This is not multi-threading; it only lets Qt repaint the window first.
        runButton->setEnabled(false);
        logEdit->setPlainText("執行中，請稍候...");
        QApplication::processEvents();

        QString log;
        QString error;
        const bool ok = processWorkbook(inputPath, outputPath, &log, &error);

        logEdit->setPlainText(log.trimmed());
        runButton->setEnabled(true);

        if (ok) {
            QMessageBox::information(&window,
                                     "完成",
                                     "平均資料已輸出：\n" + QDir::toNativeSeparators(outputPath));
        } else {
            logEdit->appendPlainText("\nError: " + error);
            QMessageBox::critical(&window, "失敗", error);
        }
    });

    window.show();
    return app.exec();
}
