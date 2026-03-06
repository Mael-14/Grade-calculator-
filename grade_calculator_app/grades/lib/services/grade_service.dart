import 'dart:io';
import 'dart:typed_data';
import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';

/// A single student record with name, score, and calculated grade.
class StudentRecord {
  final String name;
  final double score;
  final String grade;

  const StudentRecord({
    required this.name,
    required this.score,
    required this.grade,
  });
}

/// Result summary returned after processing an Excel file.
class GradingResult {
  final String outputPath;
  final int totalStudents;
  final List<StudentRecord> students;
  final Map<String, int> gradeDistribution;

  const GradingResult({
    required this.outputPath,
    required this.totalStudents,
    required this.students,
    required this.gradeDistribution,
  });
}

/// Core grading service using higher-order functions and lambdas
/// for concise, functional-style processing.
class GradeService {
  // ── Grade scale: ordered list of (minScore, grade) ────────────────
  static const List<MapEntry<double, String>> _scale = [
    MapEntry(90, 'A'),
    MapEntry(85, 'B+'),
    MapEntry(80, 'B'),
    MapEntry(75, 'C+'),
    MapEntry(70, 'C'),
    MapEntry(65, 'D+'),
    MapEntry(60, 'D'),
  ];

  /// Convert a numeric score to a letter grade via firstWhere lambda.
  static String calculateGrade(double score) =>
      _scale
          .cast<MapEntry<double, String>?>()
          .firstWhere((e) => score >= e!.key, orElse: () => null)
          ?.value ??
      'F';

  /// Map a grade letter to its display color.
  static Color gradeColor(String grade) =>
      const <String, Color>{
        'A': Color(0xFF4CAF50),
        'B+': Color(0xFF66BB6A),
        'B': Color(0xFF8BC34A),
        'C+': Color(0xFFFFC107),
        'C': Color(0xFFFF9800),
        'D+': Color(0xFFFF5722),
        'D': Color(0xFFF44336),
        'F': Color(0xFFB71C1C),
      }[grade] ??
      Colors.grey;

  // ── Private helpers  ──────────────────────────────────────────────

  /// Extract a cell's value as a plain string.
  static String _cellToString(CellValue? cell) => switch (cell) {
    IntCellValue c => c.value.toString(),
    DoubleCellValue c => c.value.toString(),
    TextCellValue c => c.value.toString(),
    _ => '',
  };

  /// Try to parse a cell as a double score.
  static double? _cellToDouble(Data? cell) => switch (cell?.value) {
    IntCellValue c => c.value.toDouble(),
    DoubleCellValue c => c.value,
    _ =>
      cell?.value != null ? double.tryParse(_cellToString(cell!.value)) : null,
  };

  /// Detect the column index whose header matches any of the given keywords
  /// within the first [maxScan] rows.  Returns `(headerRowIndex, colIndex)`.
  static (int, int) _detectColumn(
    Sheet sheet, {
    List<String> keywords = const ['mark', 'score', 'grade', 'points'],
    int maxScan = 5,
  }) {
    for (var r = 0; r < sheet.maxRows && r < maxScan; r++) {
      final row = sheet.row(r);
      final idx = row.indexWhere(
        (c) =>
            keywords.contains(c?.value?.toString().toLowerCase().trim() ?? ''),
      );
      if (idx != -1) return (r, idx);
    }
    return (0, -1);
  }

  // ── Public API  ───────────────────────────────────────────────────

  /// Process an Excel file [bytes] and return a [GradingResult].
  ///
  /// Uses higher-order `map` / `fold` to build student list and
  /// grade distribution in a single pass.
  static Future<GradingResult> processExcelBytes(
    Uint8List bytes, {
    required String originalFileName,
  }) async {
    final excel = Excel.decodeBytes(bytes);

    // Find the first non-empty sheet
    final inputSheet = excel.tables.values.firstWhere(
      (s) => s.maxRows > 0,
      orElse: () => throw Exception('No data found in the uploaded file.'),
    );

    // Detect header & score column
    final (headerRow, scoreCol) = _detectColumn(inputSheet);
    if (scoreCol == -1) {
      throw Exception(
        'Could not find a "Mark", "Score", "Grade", or "Points" column.\n'
        'Make sure your Excel header row contains one of those names.',
      );
    }

    // Detect optional name column
    final nameCol = inputSheet
        .row(headerRow)
        .indexWhere(
          (c) => [
            'name',
            'student',
            'student name',
            'student_name',
          ].contains(c?.value?.toString().toLowerCase().trim() ?? ''),
        );

    // ── Build student list using map + where ──────────────────────
    final students =
        List.generate(
              inputSheet.maxRows - headerRow - 1,
              (i) => inputSheet.row(headerRow + 1 + i),
            )
            .where((row) => row.isNotEmpty)
            .map((row) {
              final score = scoreCol < row.length
                  ? _cellToDouble(row[scoreCol])
                  : null;
              final name = nameCol != -1 && nameCol < row.length
                  ? _cellToString(row[nameCol]?.value)
                  : 'Student';
              return score != null
                  ? StudentRecord(
                      name: name.isEmpty ? 'Student' : name,
                      score: score,
                      grade: calculateGrade(score),
                    )
                  : null;
            })
            .whereType<StudentRecord>()
            .toList();

    // ── Grade distribution via fold ──────────────────────────────
    final distribution = students.fold<Map<String, int>>(
      {},
      (map, s) => map..update(s.grade, (v) => v + 1, ifAbsent: () => 1),
    );

    // ── Build output Excel ──────────────────────────────────────
    final output = Excel.createExcel();
    output.rename('Sheet1', 'Grades');
    final outSheet = output['Grades'];

    // Header row — copy originals + append "Calculated Grade"
    final originalHeaders =
        inputSheet
            .row(headerRow)
            .map((c) => TextCellValue(_cellToString(c?.value)))
            .toList()
          ..add(TextCellValue('Calculated Grade'));
    outSheet.appendRow(originalHeaders);

    // Data rows
    for (var i = headerRow + 1; i < inputSheet.maxRows; i++) {
      final row = inputSheet.row(i);
      if (row.isEmpty) continue;

      final cells = row.map<CellValue>((c) {
        if (c?.value == null) return TextCellValue('');
        return switch (c!.value) {
          IntCellValue v => IntCellValue(v.value),
          DoubleCellValue v => DoubleCellValue(v.value),
          _ => TextCellValue(_cellToString(c.value)),
        };
      }).toList();

      final score = scoreCol < row.length ? _cellToDouble(row[scoreCol]) : null;
      cells.add(
        TextCellValue(
          score != null ? calculateGrade(score) : 'Invalid/Missing',
        ),
      );
      outSheet.appendRow(cells);
    }

    // Save to temp directory
    final dir = await getApplicationDocumentsDirectory();
    final outName = 'graded_${originalFileName}';
    final outPath = p.join(dir.path, outName);
    final fileBytes = output.save();
    if (fileBytes == null) throw Exception('Failed to encode output Excel.');
    await File(outPath).writeAsBytes(fileBytes);

    return GradingResult(
      outputPath: outPath,
      totalStudents: students.length,
      students: students,
      gradeDistribution: distribution,
    );
  }

  /// Create a sample/template Excel file and return its path.
  static Future<String> createTemplateFile() async {
    final excel = Excel.createExcel();
    final sheet = excel['Sheet1'];

    sheet.appendRow([TextCellValue('Student Name'), TextCellValue('Score')]);

    // Sample data using map + forEach
    [
          ('Alice Johnson', 95),
          ('Bob Smith', 88),
          ('Charlie Brown', 72),
          ('Diana Prince', 60),
          ('Eve Davis', 45),
        ]
        .map((e) => [TextCellValue(e.$1), IntCellValue(e.$2)])
        .forEach(sheet.appendRow);

    final dir = await getApplicationDocumentsDirectory();
    final path = p.join(dir.path, 'grade_template.xlsx');
    final bytes = excel.save();
    if (bytes != null) await File(path).writeAsBytes(bytes);
    return path;
  }
}
