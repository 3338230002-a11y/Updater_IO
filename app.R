### =============================================================================
### APLIKASI PEMBARUAN & ANALISIS TABEL INPUT-OUTPUT (I-O) — IPFP
### VERSI FINAL TERVERIFIKASI (MERGED + LEONTIEF/GHOSH + GRAFIK FPD-BPD)
### =============================================================================
library(shiny)
library(bslib)
library(readxl)
library(openxlsx)
library(dplyr)
library(DT)
library(plotly)
library(ggplot2)
options(shiny.maxRequestLimit = 50 * 1024^2)
options(shiny.maxRequestSize = 50 * 1024^2)

### =============================================================================
### KONSTANTA: KODE & NAMA SEKTOR
### =============================================================================
CODES_52 <- sprintf("I-%02d", 1:52)
NAMES_52 <- c(
  "Pertanian Tanaman Pangan",
  "Pertanian Tanaman Hortikultura Semusim, Hortikultura Tahunan, dan Lainnya",
  "Perkebunan Semusim dan Tahunan",
  "Peternakan",
  "Jasa Pertanian dan Perburuan",
  "Kehutanan dan Penebangan Kayu",
  "Perikanan",
  "Pertambangan Minyak, Gas dan Panas Bumi",
  "Pertambangan Batubara dan Lignit",
  "Pertambangan Bijih Logam",
  "Pertambangan dan Penggalian Lainnya",
  "Industri Batubara dan Pengilangan Migas",
  "Industri Makanan dan Minuman",
  "Industri Pengolahan Tembakau",
  "Industri Tekstil dan Pakaian Jadi",
  "Industri Kulit, Barang dari Kulit dan Alas Kaki",
  "Industri Kayu, Barang dari Kayu dan Gabus dan Barang Anyaman dari Bambu, Rotan dan Sejenisnya",
  "Industri Kertas dan Barang dari Kertas, Percetakan dan Reproduksi Media Rekaman",
  "Industri Kimia, Farmasi dan Obat Tradisional",
  "Industri Karet, Barang dari Karet dan Plastik",
  "Industri Barang Galian bukan Logam",
  "Industri Logam Dasar",
  "Industri Barang dari Logam, Komputer, Barang Elektronik, Optik dan Peralatan Listrik",
  "Industri Mesin dan Perlengkapan YTDL",
  "Industri Alat Angkutan",
  "Industri Furnitur",
  "Industri Pengolahan Lainnya, Jasa Reparasi dan Pemasangan Mesin dan Peralatan",
  "Ketenagalistrikan",
  "Pengadaan Gas dan Produksi Es",
  "Pengadaan Air, Pengelolaan Sampah, Limbah, dan Daur Ulang",
  "Konstruksi",
  "Perdagangan Mobil, Sepeda Motor dan Reparasinya",
  "Perdagangan Besar dan Eceran, Bukan Mobil dan Sepeda Motor",
  "Angkutan Rel",
  "Angkutan Darat",
  "Angkutan Laut",
  "Angkutan Sungai Danau dan Penyeberangan",
  "Angkutan Udara",
  "Pergudangan dan Jasa Penunjang Angkutan, Pos dan Kurir",
  "Penyediaan Akomodasi",
  "Penyediaan Makan Minum",
  "Jasa Informasi dan Komunikasi",
  "Jasa Perantara Keuangan Selain Bank Sentral",
  "Asuransi dan Dana Pensiun",
  "Jasa Keuangan Lainnya",
  "Jasa Penunjang Keuangan",
  "Real Estate",
  "Jasa Perusahaan",
  "Administrasi Pemerintahan, Pertahanan dan Jaminan Sosial Wajib",
  "Jasa Pendidikan",
  "Jasa Kesehatan dan Kegiatan Sosial",
  "Jasa Lainnya"
)

CODES_17 <- c(
  "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
  "K", "L", "M,N", "O", "P", "Q", "R,S,T,U"
)
NAMES_17 <- c(
  "Pertanian, Kehutanan, dan Perikanan",
  "Pertambangan dan Penggalian",
  "Industri Pengolahan",
  "Pengadaan Listrik dan Gas",
  "Pengadaan Air, Pengelolaan Sampah, Limbah dan Daur Ulang",
  "Konstruksi",
  "Perdagangan Besar dan Eceran; Reparasi Mobil dan Sepeda Motor",
  "Transportasi dan Pergudangan",
  "Penyediaan Akomodasi dan Makan Minum",
  "Informasi dan Komunikasi",
  "Jasa Keuangan dan Asuransi",
  "Real Estate",
  "Jasa Perusahaan",
  "Administrasi Pemerintahan, Pertahanan dan Jaminan Sosial Wajib",
  "Jasa Pendidikan",
  "Jasa Kesehatan dan Kegiatan Sosial",
  "Jasa lainnya"
)

LABOR_HEADER_TITLE <- "Penduduk Usia 15 Tahun Ke Atas Yang Bekerja Menurut Lapangan Pekerjaan Utama"

### =============================================================================
### PARSER NUMERIK ROBUST
### =============================================================================
clean_num <- function(x) {
  vals <- as.character(unlist(x))
  vapply(vals, function(s) {
    s <- trimws(s)
    if (is.na(s) || s %in% c("", "-", "--", ".", "NA", "#N/A", "#REF!", "#DIV/0!", "#VALUE!", "#NUM!", "#NULL!")) return(0)
    neg <- grepl("^\\(.*\\)$", s)
    s <- gsub("[()]", "", s)
    s <- gsub("\\s", "", s)
    if (grepl("^\\d{1,3}(,\\d{3})+(\\.\\d+)?$", s)) {
      s <- gsub(",", "", s)
    } else if (grepl("^\\d{1,3}(\\.\\d{3})+(,\\d+)?$", s)) {
      s <- gsub("\\.", "", s)
      s <- sub(",", ".", s)
    } else if (grepl("^\\d+,\\d+$", s)) {
      s <- sub(",", ".", s)
    } else {
      s <- gsub(",", "", s)
    }
    v <- suppressWarnings(as.numeric(s))
    if (is.na(v) || !is.finite(v)) return(0)
    if (neg) v <- -v
    v
  }, numeric(1), USE.NAMES = FALSE)
}

### =============================================================================
### FUNGSI BANTU PEMBACAAN & DETEKSI STRUKTUR
### =============================================================================
read_raw_text_matrix <- function(path, sheet) {
  df <- readxl::read_excel(
    path,
    sheet = sheet,
    col_names = FALSE,
    col_types = "text",
    .name_repair = "minimal"
  )
  as.matrix(df)
}

find_row_index_in_col <- function(mat, col_idx, target) {
  if (col_idx > ncol(mat)) return(NA_integer_)
  col_vals <- trimws(as.character(mat[, col_idx]))
  idx <- which(col_vals == target)
  if (length(idx) == 0) return(NA_integer_)
  idx[1]
}

find_col_index_in_row <- function(mat, row_idx, target) {
  if (is.na(row_idx) || row_idx < 1 || row_idx > nrow(mat)) return(NA_integer_)
  row_vals <- trimws(as.character(mat[row_idx, ]))
  idx <- which(row_vals == target)
  if (length(idx) == 0) return(NA_integer_)
  idx[1]
}

find_col_index_in_row_ci <- function(mat, row_idx, target) {
  if (is.na(row_idx) || row_idx < 1 || row_idx > nrow(mat)) return(NA_integer_)
  row_vals <- toupper(trimws(as.character(mat[row_idx, ])))
  idx <- which(row_vals == toupper(target))
  if (length(idx) == 0) return(NA_integer_)
  idx[1]
}

find_row_with_all <- function(mat, targets) {
  for (i in seq_len(nrow(mat))) {
    v <- trimws(as.character(mat[i, ]))
    if (all(targets %in% v)) return(i)
  }
  NA_integer_
}

find_row_with_all_ci <- function(mat, targets) {
  tg <- toupper(targets)
  for (i in seq_len(nrow(mat))) {
    v <- toupper(trimws(as.character(mat[i, ])))
    if (all(tg %in% v)) return(i)
  }
  NA_integer_
}

find_code_col_base <- function(mat, max_scan = 6) {
  known <- c(CODES_52, CODES_17, "2000", "2001", "2090", "2100")
  for (cc in seq_len(min(max_scan, ncol(mat)))) {
    v <- trimws(as.character(mat[, cc]))
    if (any(v %in% known)) return(cc)
  }
  NA_integer_
}

detect_sheet_type <- function(path, sheet) {
  mat <- tryCatch(read_raw_text_matrix(path, sheet), error = function(e) NULL)
  if (is.null(mat) || nrow(mat) < 2 || ncol(mat) < 2) return(NULL)
  has_io_header <- !is.na(find_row_with_all(mat, c("3090", "3100")))
  has_sector_code <- FALSE
  for (cc in seq_len(min(6, ncol(mat)))) {
    vals <- trimws(as.character(mat[, cc]))
    if (any(vals %in% CODES_52) || any(vals %in% CODES_17)) {
      has_sector_code <- TRUE
      break
    }
  }
  if (has_io_header && has_sector_code) return("base")
  if (!is.na(find_row_with_all_ci(mat, c("Total Input", "Total Output")))) return("target")
  NULL
}

resolve_inputs <- function(file_list) {
  base_loc <- NULL
  target_loc <- NULL
  for (f in file_list) {
    if (is.null(f)) next
    path <- f$datapath
    sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character(0))
    for (sh in sheets) {
      type <- detect_sheet_type(path, sh)
      if (identical(type, "base") && is.null(base_loc)) {
        base_loc <- list(path = path, sheet = sh)
      } else if (identical(type, "target") && is.null(target_loc)) {
        target_loc <- list(path = path, sheet = sh)
      }
    }
  }
  if (is.null(base_loc)) {
    stop("Tabel I-O Data Dasar tidak terdeteksi. Gunakan Template Input Matriks 52x52 atau 17x17.")
  }
  if (is.null(target_loc)) {
    stop("Data Target tidak terdeteksi. Gunakan template Data Target yang telah disediakan.")
  }
  list(
    base_path = base_loc$path,
    base_sheet = base_loc$sheet,
    target_path = target_loc$path,
    target_sheet = target_loc$sheet
  )
}

find_base_sheet_in_file <- function(path) {
  sheets <- tryCatch(readxl::excel_sheets(path), error = function(e) character(0))
  for (sh in sheets) {
    type <- tryCatch(detect_sheet_type(path, sh), error = function(e) NULL)
    if (identical(type, "base")) {
      return(list(path = path, sheet = sh))
    }
  }
  stop(
    "Sheet Tabel I-O untuk analisis tidak ditemukan. ",
    "Gunakan Template Input Matriks 52x52 atau 17x17."
  )
}

read_base_io <- function(path, sheet) {
  mat <- read_raw_text_matrix(path, sheet)
  code_col <- find_code_col_base(mat)
  if (is.na(code_col)) stop("Struktur Tabel I-O tidak dikenali: kolom Kode sektor tidak ditemukan.")
  header_row <- find_row_with_all(mat, c("3090", "3100"))
  if (is.na(header_row)) stop("Baris kode (3090/3100) tidak ditemukan pada tabel I-O.")
  
  row_impor_ln    <- find_row_index_in_col(mat, code_col, "2000")
  row_impor_ap    <- find_row_index_in_col(mat, code_col, "2001")
  row_ntb         <- find_row_index_in_col(mat, code_col, "2090")
  row_total_input <- find_row_index_in_col(mat, code_col, "2100")
  
  required_codes <- c(
    "2000" = "Input Antara Impor Luar Negeri",
    "2001" = "Input Antara Impor Antar Provinsi",
    "2090" = "Nilai Tambah Bruto",
    "2100" = "Total Input"
  )
  required_rows <- c(row_impor_ln, row_impor_ap, row_ntb, row_total_input)
  if (any(is.na(required_rows))) {
    missing <- names(required_codes)[is.na(required_rows)]
    msg <- paste0(
      "Kode ", missing, " - ",
      unname(required_codes[missing]),
      " belum tersedia."
    )
    stop(paste(msg, collapse = " | "))
  }
  
  col_vals <- trimws(as.character(mat[, code_col]))
  agg_rows <- which(col_vals %in% c("2000", "2001", "2090", "2100"))
  agg_rows <- agg_rows[agg_rows > header_row]
  if (length(agg_rows) == 0) stop("Baris agregat tidak ditemukan setelah baris header.")
  first_agg <- min(agg_rows)
  start_row <- header_row + 1
  end_row <- first_agg - 1
  if (end_row < start_row) stop("Tidak ditemukan baris sektor di antara header dan baris agregat.")
  n <- end_row - start_row + 1
  
  codes <- trimws(as.character(mat[start_row:end_row, code_col]))
  name_col <- if (code_col > 1) code_col - 1 else code_col + 1
  names_sektor <- trimws(as.character(mat[start_row:end_row, name_col]))
  names_sektor[is.na(names_sektor) | names_sektor == ""] <-
    codes[is.na(names_sektor) | names_sektor == ""]
  
  mcols <- vapply(
    codes,
    function(cd) find_col_index_in_row(mat, header_row, cd),
    integer(1),
    USE.NAMES = FALSE
  )
  if (any(is.na(mcols))) {
    sc <- find_col_index_in_row(mat, header_row, codes[1])
    if (is.na(sc)) stop("Kolom matriks transaksi tidak dapat ditentukan.")
    mcols <- seq(sc, sc + n - 1)
  }
  start_col <- mcols[1]
  end_col <- mcols[n]
  if (end_col > ncol(mat)) stop("Jumlah kolom matriks kurang dari jumlah sektor.")
  
  Z <- matrix(
    clean_num(mat[start_row:end_row, start_col:end_col, drop = FALSE]),
    nrow = n,
    ncol = n
  )
  
  col_konsumsi <- find_col_index_in_row(mat, header_row, "3090")
  col_output   <- find_col_index_in_row(mat, header_row, "3100")
  if (is.na(col_konsumsi) || is.na(col_output)) stop("Kolom 3090/3100 tidak ditemukan.")
  
  konsumsi_akhir    <- clean_num(mat[start_row:end_row, col_konsumsi])
  total_output_base <- clean_num(mat[start_row:end_row, col_output])
  impor_ln <- clean_num(mat[row_impor_ln, start_col:end_col])
  impor_ap <- clean_num(mat[row_impor_ap, start_col:end_col])
  ntb      <- clean_num(mat[row_ntb, start_col:end_col])
  total_input_base <- clean_num(mat[row_total_input, start_col:end_col])
  
  fd <- list()
  for (fc in c("3011", "3012", "3020", "3030", "3041", "3071", "3072")) {
    cc <- find_col_index_in_row(mat, header_row, fc)
    fd[[fc]] <- if (!is.na(cc)) clean_num(mat[start_row:end_row, cc]) else NULL
  }
  
  list(
    codes = codes,
    names = names_sektor,
    Z = Z,
    konsumsi_akhir = konsumsi_akhir,
    total_output = total_output_base,
    impor_ln = impor_ln,
    impor_ap = impor_ap,
    ntb = ntb,
    total_input = total_input_base,
    n = n,
    fd = fd
  )
}

read_target_io <- function(path, sheet) {
  mat <- read_raw_text_matrix(path, sheet)
  header_row <- find_row_with_all_ci(mat, c("Total Input", "Total Output"))
  if (is.na(header_row)) {
    stop("Struktur Data Tahun Target tidak dikenali: header 'Total Input' dan 'Total Output' tidak ditemukan.")
  }
  code_col <- find_col_index_in_row_ci(mat, header_row, "Kode")
  if (is.na(code_col)) code_col <- 1
  col_ti <- find_col_index_in_row_ci(mat, header_row, "Total Input")
  col_to <- find_col_index_in_row_ci(mat, header_row, "Total Output")
  if (is.na(col_ti) || is.na(col_to)) {
    stop("Kolom Total Input/Total Output pada Data Tahun Target tidak ditemukan.")
  }
  codes <- character(0)
  rows <- integer(0)
  for (i in seq.int(header_row + 1, nrow(mat))) {
    cd <- trimws(as.character(mat[i, code_col]))
    if (is.na(cd) || cd == "") break
    codes <- c(codes, cd)
    rows <- c(rows, i)
  }
  codes[codes %in% c("MN", "M-N")] <- "M,N"
  codes[codes %in% c("RSTU", "R-S-T-U")] <- "R,S,T,U"
  n <- length(codes)
  if (n == 0) stop("Tidak ditemukan baris sektor pada Data Tahun Target.")
  is52 <- n == 52 && all(codes == CODES_52)
  is17 <- n == 17 && all(codes == CODES_17)
  if (!is52 && !is17) {
    stop("Data Tahun Target harus berisi tepat 52 sektor (I-01 s.d. I-52) atau 17 sektor (A-L, M,N, O, P, Q, R,S,T,U).")
  }
  name_col <- find_col_index_in_row_ci(mat, header_row, "Nama Sektor")
  names_sektor <- if (!is.na(name_col)) {
    trimws(as.character(mat[rows, name_col]))
  } else {
    codes
  }
  list(
    codes = codes,
    names = names_sektor,
    total_input = clean_num(mat[rows, col_ti]),
    total_output = clean_num(mat[rows, col_to]),
    n = n,
    version = if (is52) 52 else 17
  )
}

align_target_to_base <- function(base_codes, target) {
  base_is52 <- length(base_codes) == 52 && all(base_codes == CODES_52)
  base_is17 <- length(base_codes) == 17 && all(base_codes == CODES_17)
  if (base_is52 && target$version != 52) {
    stop("Data Dasar berukuran 52x52, sehingga Data Tahun Target juga harus menggunakan template 52 sektor.")
  }
  if (base_is17 && target$version != 17) {
    stop("Data Dasar berukuran 17x17, sehingga Data Tahun Target juga harus menggunakan template 17 sektor.")
  }
  idx <- match(base_codes, target$codes)
  if (any(is.na(idx))) {
    stop(paste0(
      "Kode sektor tidak ditemukan pada Data Tahun Target: ",
      paste(base_codes[is.na(idx)], collapse = ", ")
    ))
  }
  list(
    total_input = target$total_input[idx],
    total_output = target$total_output[idx]
  )
}

make_analysis_object_from_base <- function(base, source_type) {
  list(
    codes = base$codes,
    names = base$names,
    res = list(
      Z_new = base$Z,
      iterations = 0,
      converged = TRUE,
      max_diff = 0,
      konsumsi_akhir_new = base$konsumsi_akhir,
      impor_ln_new = base$impor_ln,
      impor_ap_new = base$impor_ap,
      ntb_new = base$ntb,
      total_output_new = base$total_output,
      total_input_new = base$total_input,
      max_input_deviation = 0,
      scale_factor = 1,
      neg_u_count = 0,
      neg_v_count = 0
    ),
    fd = base$fd,
    base_output = base$total_output,
    error = NULL,
    source = source_type
  )
}

validate_target_dimension <- function(base, target) {
  base_n <- length(base$codes)
  target_n <- length(target$codes)
  if (!(base_n %in% c(52, 17))) {
    stop("Ukuran Data Dasar tidak didukung. Gunakan matriks 52x52 atau 17x17.")
  }
  if (!(target_n %in% c(52, 17))) {
    stop("Ukuran Data Tahun Target tidak didukung. Gunakan matriks 52x52 atau 17x17.")
  }
  if (base_n == 52 && target_n == 17) {
    stop(paste0(
      "Ukuran Data Tahun Target tidak sesuai. ",
      "Data Dasar yang diunggah berukuran 52x52, sehingga Data Tahun Target wajib menggunakan matriks 52x52. ",
      "Silakan unggah Template Tahun Target - 52x52.xlsx."
    ))
  }
  if (base_n == 17 && target_n == 52) {
    stop(paste0(
      "Ukuran Data Tahun Target tidak sesuai. ",
      "Data Dasar yang diunggah berukuran 17x17, sehingga Data Tahun Target wajib menggunakan matriks 17x17. ",
      "Silakan unggah Template Tahun Target - 17x17.xlsx."
    ))
  }
  if (base_n != target_n) {
    stop(sprintf(
      "Ukuran Data Tahun Target tidak sesuai. Data Dasar memiliki %d sektor, sedangkan Data Tahun Target memiliki %d sektor.",
      base_n, target_n
    ))
  }
  invisible(TRUE)
}

### =============================================================================
### AGREGASI 52 -> 17
### =============================================================================
map52to17 <- function(cd) {
  num <- suppressWarnings(as.integer(sub("^I-", "", cd)))
  if (is.na(num)) return(NA_character_)
  if (num >= 1 && num <= 7) return("A")
  if (num <= 11) return("B")
  if (num <= 27) return("C")
  if (num <= 29) return("D")
  if (num == 30) return("E")
  if (num == 31) return("F")
  if (num <= 33) return("G")
  if (num <= 39) return("H")
  if (num <= 41) return("I")
  if (num == 42) return("J")
  if (num <= 46) return("K")
  if (num == 47) return("L")
  if (num == 48) return("M,N")
  if (num == 49) return("O")
  if (num == 50) return("P")
  if (num == 51) return("Q")
  if (num == 52) return("R,S,T,U")
  NA_character_
}

make_aggregation <- function(codes52) {
  grp <- vapply(codes52, map52to17, character(1), USE.NAMES = FALSE)
  if (any(is.na(grp))) return(NULL)
  G <- t(vapply(CODES_17, function(k) as.integer(grp == k), integer(length(codes52))))
  list(G = G, grp = grp)
}

### =============================================================================
### DETEKSI DATA TENAGA KERJA
### =============================================================================
num_like <- function(s) {
  s <- trimws(as.character(s))
  grepl("^-?[0-9][0-9.,]*$", s) |
    grepl("^\\([0-9.,]+\\)$", s) |
    grepl("^-?[0-9]+(\\.[0-9]+)?([eE][+-]?[0-9]+)?$", s)
}

LABOR_KEYWORDS <- list(
  list("TENAGA KERJA", 4),
  list("BEKERJA", 4),
  list("PEKERJA", 4),
  list("PENDUDUK", 3),
  list("ANGKATAN KERJA", 3),
  list("LAPANGAN PEKERJAAN", 3),
  list("BURUH", 3),
  list("KARYAWAN", 3),
  list("ORANG", 2),
  list("JUMLAH", 1)
)

keyword_score <- function(s) {
  s <- toupper(trimws(as.character(s)))
  if (is.na(s) || s == "") return(0)
  sc <- 0
  for (kw in LABOR_KEYWORDS) {
    if (grepl(kw[[1]], s, fixed = TRUE)) sc <- sc + kw[[2]]
  }
  sc
}

norm_str <- function(s) {
  gsub("[^A-Z0-9]", "", toupper(trimws(as.character(s))))
}

build_sector_matcher <- function(tbl_codes) {
  is17 <- length(tbl_codes) == 17 && all(tbl_codes == CODES_17)
  is52 <- length(tbl_codes) == 52
  ncodes <- norm_str(tbl_codes)
  n17 <- norm_str(NAMES_17)
  n52 <- norm_str(NAMES_52)
  function(id) {
    id <- trimws(as.character(id))
    if (is.na(id) || id == "") return(NA_integer_)
    i <- match(id, tbl_codes)
    if (!is.na(i)) return(i)
    nid <- norm_str(id)
    if (nid == "") return(NA_integer_)
    i <- match(nid, ncodes)
    if (!is.na(i)) return(i)
    if (is17) {
      g <- map52to17(id)
      if (!is.na(g)) return(match(g, tbl_codes))
    }
    if (is52) {
      i <- match(id, CODES_52)
      if (!is.na(i)) return(i)
    }
    i <- match(nid, n17)
    if (!is.na(i)) return(if (is17) i else NA_integer_)
    i <- match(nid, n52)
    if (!is.na(i)) {
      if (is52) return(i)
      if (is17) return(match(map52to17(CODES_52[i]), tbl_codes))
    }
    if (nchar(nid) >= 12) {
      for (k in seq_along(n17)) {
        if (nchar(n17[k]) >= 12 &&
            (grepl(n17[k], nid, fixed = TRUE) || grepl(nid, n17[k], fixed = TRUE))) {
          return(if (is17) k else NA_integer_)
        }
      }
      for (k in seq_along(n52)) {
        if (nchar(n52[k]) >= 12 &&
            (grepl(n52[k], nid, fixed = TRUE) || grepl(nid, n52[k], fixed = TRUE))) {
          if (is52) return(k)
          if (is17) return(match(map52to17(CODES_52[k]), tbl_codes))
        }
      }
    }
    NA_integer_
  }
}

extract_labor_candidate <- function(mat, tbl_codes, matcher) {
  if (is.null(mat) || nrow(mat) < 3 || ncol(mat) < 2) return(NULL)
  n <- length(tbl_codes)
  scan_rows <- seq_len(min(25, nrow(mat)))
  key_score <- vapply(seq_len(ncol(mat)), function(cc) {
    max(c(0, vapply(as.character(mat[scan_rows, cc]), keyword_score, numeric(1))))
  }, numeric(1))
  header_of <- function(cc, upto) {
    if (is.na(upto) || upto < 1) return("")
    rr <- seq_len(min(25, upto))
    txt <- trimws(as.character(mat[rr, cc]))
    txt <- txt[!is.na(txt) & txt != ""]
    if (length(txt) == 0) return("")
    txt[length(txt)]
  }
  id_col <- NA_integer_
  id_count <- 0
  for (cc in seq_len(ncol(mat))) {
    hits <- sum(!is.na(vapply(as.character(mat[, cc]), matcher, integer(1))))
    if (hits > id_count) {
      id_count <- hits
      id_col <- cc
    }
  }
  if (!is.na(id_col) && id_count >= max(3, ceiling(n / 2))) {
    mrows <- vapply(as.character(mat[, id_col]), matcher, integer(1))
    data_rows <- which(!is.na(mrows))
    idx_rows <- mrows[data_rows]
    best_col <- NA_integer_
    best_score <- -Inf
    best_cov <- 0
    for (cc in seq_len(ncol(mat))) {
      if (cc == id_col) next
      raw <- as.character(mat[data_rows, cc])
      cov <- mean(num_like(raw))
      sc <- cov * 100 + key_score[cc]
      if (cov >= 0.6 && sc > best_score) {
        best_score <- sc
        best_col <- cc
        best_cov <- cov
      }
    }
    if (is.na(best_col)) return(NULL)
    raw <- as.character(mat[data_rows, best_col])
    ok <- num_like(raw)
    ok_rows <- which(ok)
    if (length(ok_rows) == 0) return(NULL)
    vals <- clean_num(raw[ok_rows])
    agg <- tapply(vals, idx_rows[ok_rows], sum)
    out <- rep(NA_real_, n)
    pos <- as.integer(names(agg))
    out[pos] <- as.numeric(agg)
    if (any(is.na(out))) return(NULL)
    list(
      values = out,
      score = 1000 + best_cov * 100 + key_score[best_col],
      info = list(
        method = "pencocokan kode/nama sektor",
        column = header_of(best_col, min(data_rows) - 1),
        col_index = best_col,
        matched = n
      )
    )
  } else {
    row_scores <- vapply(scan_rows, function(r) {
      max(c(0, vapply(as.character(mat[r, ]), keyword_score, numeric(1))))
    }, numeric(1))
    hdr_row <- if (max(row_scores) > 0) which.max(row_scores) else 0
    rows_b <- seq_len(nrow(mat))
    if (hdr_row > 0) rows_b <- rows_b[rows_b > hdr_row]
    best_col <- NA_integer_
    best_score <- -Inf
    for (cc in seq_len(ncol(mat))) {
      dens <- mean(num_like(as.character(mat[rows_b, cc])))
      sc <- dens * 100 + key_score[cc]
      if (sc > best_score) {
        best_score <- sc
        best_col <- cc
      }
    }
    raw <- as.character(mat[, best_col])
    vals <- clean_num(raw[num_like(raw)])
    if (length(vals) < n) return(NULL)
    list(
      values = vals[seq_len(n)],
      score = key_score[best_col],
      info = list(
        method = "urutan baris sesuai tabel",
        column = header_of(best_col, hdr_row),
        col_index = best_col,
        matched = n
      )
    )
  }
}

read_labor_data_multi <- function(paths, tbl_codes) {
  matcher <- build_sector_matcher(tbl_codes)
  best <- NULL
  for (p in paths) {
    if (is.null(p)) next
    sheets <- tryCatch(readxl::excel_sheets(p), error = function(e) character(0))
    for (sh in sheets) {
      mat <- tryCatch(read_raw_text_matrix(p, sh), error = function(e) NULL)
      cand <- extract_labor_candidate(mat, tbl_codes, matcher)
      if (!is.null(cand)) {
        cand$info$file <- basename(p)
        cand$info$sheet <- sh
        if (is.null(best) || cand$score > best$score) best <- cand
      }
    }
  }
  if (is.null(best)) {
    stop(paste0(
      "Data pekerja tidak terdeteksi pada seluruh berkas/sheet yang diunggah. ",
      "Untuk analisis 52 sektor gunakan kode I-01 s.d. I-52; untuk analisis 17 sektor gunakan kode A-L, M,N, O, P, Q, R,S,T,U. ",
      "Gunakan template yang sesuai pada menu Unduh Template."
    ))
  }
  if (length(tbl_codes) == 52 && best$info$matched != 52) {
    stop("Data pekerja tidak sesuai dengan analisis 52 sektor. Unggah data dengan kode I-01 s.d. I-52.")
  }
  if (length(tbl_codes) == 17 && best$info$matched != 17) {
    stop("Data pekerja tidak sesuai dengan analisis 17 sektor. Unggah data dengan kode A-L, M,N, O, P, Q, R,S,T,U.")
  }
  best
}

### =============================================================================
### IPFP
### =============================================================================
run_ipfp <- function(Z0, u_star, v_star, max_iter = 1000, tol = 1e-6,
                     session = NULL, progress_start = 60, progress_end = 90) {
  Z <- Z0
  max_diff <- NA_real_
  iter_used <- 0
  converged <- FALSE
  report_every <- max(1, floor(max_iter / 30))
  for (iter in seq_len(max_iter)) {
    row_sums <- rowSums(Z)
    r <- ifelse(row_sums > 0, u_star / row_sums, 0)
    r[!is.finite(r)] <- 0
    Z <- sweep(Z, 1, r, "*")
    col_sums <- colSums(Z)
    s <- ifelse(col_sums > 0, v_star / col_sums, 0)
    s[!is.finite(s)] <- 0
    Z <- sweep(Z, 2, s, "*")
    max_diff <- suppressWarnings(max(
      abs(rowSums(Z) - u_star),
      abs(colSums(Z) - v_star),
      na.rm = TRUE
    ))
    iter_used <- iter
    if (!is.na(max_diff) && is.finite(max_diff) && max_diff < tol) {
      converged <- TRUE
      break
    }
    if (!is.null(session) && iter %% report_every == 0) {
      pct <- min(progress_start + round((iter / max_iter) * (progress_end - progress_start)),
                 progress_end)
      session$sendCustomMessage(
        "updateProgress",
        list(
          percent = pct,
          message = sprintf("Iterasi IPFP ke-%d dari maks %d...", iter, max_iter)
        )
      )
      Sys.sleep(0.03)
    }
  }
  list(
    Z = Z,
    iterations = iter_used,
    converged = converged,
    max_diff = max_diff
  )
}

compute_ipfp_update <- function(base, target_aligned, max_iter, tol, session = NULL) {
  update_prog <- function(pct, msg) {
    if (!is.null(session)) {
      session$sendCustomMessage("updateProgress", list(percent = pct, message = msg))
      Sys.sleep(0.12)
    }
  }
  update_prog(35, "Menghitung rasio proporsi historis tahun dasar...")
  fd_ratio  <- ifelse(base$total_output > 0, base$konsumsi_akhir / base$total_output, 0)
  iln_ratio <- ifelse(base$total_input  > 0, base$impor_ln / base$total_input, 0)
  iap_ratio <- ifelse(base$total_input  > 0, base$impor_ap / base$total_input, 0)
  ntb_ratio <- ifelse(base$total_input  > 0, base$ntb / base$total_input, 0)
  fd_ratio[!is.finite(fd_ratio)] <- 0
  iln_ratio[!is.finite(iln_ratio)] <- 0
  iap_ratio[!is.finite(iap_ratio)] <- 0
  ntb_ratio[!is.finite(ntb_ratio)] <- 0
  update_prog(42, "Mengestimasi komponen baru berdasarkan rasio historis...")
  konsumsi_new <- fd_ratio * target_aligned$total_output
  impor_ln_new <- iln_ratio * target_aligned$total_input
  impor_ap_new <- iap_ratio * target_aligned$total_input
  ntb_new      <- ntb_ratio * target_aligned$total_input
  update_prog(50, "Menghitung margin IPFP (u* dan v*)...")
  u_star <- target_aligned$total_output - konsumsi_new
  v_star <- target_aligned$total_input - impor_ln_new - impor_ap_new - ntb_new
  neg_u_count <- sum(u_star < 0, na.rm = TRUE)
  neg_v_count <- sum(v_star < 0, na.rm = TRUE)
  u_star[u_star < 0 | is.na(u_star)] <- 0
  v_star[v_star < 0 | is.na(v_star)] <- 0
  update_prog(55, "Menyeimbangkan margin baris dan kolom...")
  sum_u <- sum(u_star)
  sum_v <- sum(v_star)
  scale_factor <- if (sum_v > 0) sum_u / sum_v else 1
  v_star_adj <- v_star * scale_factor
  update_prog(60, "Menjalankan iterasi IPFP...")
  ipfp_result <- run_ipfp(base$Z, u_star, v_star_adj, max_iter, tol, session, 60, 90)
  update_prog(92, "Menyusun hasil akhir...")
  total_output_new <- target_aligned$total_output
  total_input_new  <- colSums(ipfp_result$Z) + impor_ln_new + impor_ap_new + ntb_new
  max_input_deviation <- suppressWarnings(max(
    abs(total_input_new - target_aligned$total_input),
    na.rm = TRUE
  ))
  if (!is.finite(max_input_deviation)) max_input_deviation <- 0
  update_prog(96, "Finalisasi data...")
  list(
    Z_new = ipfp_result$Z,
    iterations = ipfp_result$iterations,
    converged = ipfp_result$converged,
    max_diff = ipfp_result$max_diff,
    konsumsi_akhir_new = konsumsi_new,
    impor_ln_new = impor_ln_new,
    impor_ap_new = impor_ap_new,
    ntb_new = ntb_new,
    total_output_new = total_output_new,
    total_input_new = total_input_new,
    max_input_deviation = max_input_deviation,
    scale_factor = scale_factor,
    neg_u_count = neg_u_count,
    neg_v_count = neg_v_count
  )
}

### =============================================================================
### TABEL ANALISIS SESUAI LEVEL SEKTOR TERPILIH (52 atau 17)
### =============================================================================
is_base_52 <- function(h) {
  std52 <- sprintf("I-%02d", 1:52)
  length(h$codes) == 52 && all(h$codes == std52)
}

build_analysis_table <- function(h, versi = "17") {
  res <- h$res
  fd_full <- list()
  for (key in names(h$fd)) {
    v <- h$fd[[key]]
    if (!is.null(v)) {
      ratio <- ifelse(h$base_output > 0, v / h$base_output, 0)
      fd_full[[key]] <- ratio * res$total_output_new
    }
  }
  source_label <- if (!is.null(h$source)) h$source else "ipfp"
  if (is_base_52(h) && identical(versi, "52")) {
    return(list(
      codes = h$codes,
      names = h$names,
      Z = res$Z_new,
      konsumsi = res$konsumsi_akhir_new,
      output = res$total_output_new,
      input = res$total_input_new,
      ntb = res$ntb_new,
      impor_ln = res$impor_ln_new,
      impor_ap = res$impor_ap_new,
      fd_new = fd_full,
      sumber = if (identical(source_label, "langsung")) {
        "Data Upload Langsung 52 sektor"
      } else {
        "Hasil IPFP 52 sektor"
      }
    ))
  }
  if (is_base_52(h)) {
    agg <- make_aggregation(h$codes)
    if (is.null(agg)) return(NULL)
    G <- agg$G
    fd_new <- lapply(fd_full, function(v) as.numeric(G %*% v))
    return(list(
      codes = CODES_17,
      names = NAMES_17,
      Z = G %*% res$Z_new %*% t(G),
      konsumsi = as.numeric(G %*% res$konsumsi_akhir_new),
      output = as.numeric(G %*% res$total_output_new),
      input = as.numeric(G %*% res$total_input_new),
      ntb = as.numeric(G %*% res$ntb_new),
      impor_ln = as.numeric(G %*% res$impor_ln_new),
      impor_ap = as.numeric(G %*% res$impor_ap_new),
      fd_new = fd_new,
      sumber = "Agregasi 52 -> 17 sektor"
    ))
  }
  if (length(h$codes) == 17) {
    return(list(
      codes = h$codes,
      names = h$names,
      Z = res$Z_new,
      konsumsi = res$konsumsi_akhir_new,
      output = res$total_output_new,
      input = res$total_input_new,
      ntb = res$ntb_new,
      impor_ln = res$impor_ln_new,
      impor_ap = res$impor_ap_new,
      fd_new = fd_full,
      sumber = if (identical(source_label, "langsung")) {
        "Data Upload Langsung 17 sektor"
      } else {
        "Hasil IPFP 17 sektor"
      }
    ))
  }
  NULL
}

### =============================================================================
### ANALISIS DAMPAK: LEONTIEF & GHOSH
### =============================================================================
compute_analysis <- function(tbl) {
  Z <- as.matrix(tbl$Z)
  n <- nrow(Z)
  X <- tbl$input
  O <- tbl$output
  NTB <- tbl$ntb
  A <- sweep(Z, 2, X, "/")
  A[!is.finite(A)] <- 0
  B <- sweep(Z, 1, O, "/")
  B[!is.finite(B)] <- 0
  I <- diag(n)
  invA <- tryCatch(solve(I - A), error = function(e) matrix(NA_real_, n, n))
  invB <- tryCatch(solve(I - B), error = function(e) matrix(NA_real_, n, n))
  okA <- !any(is.na(invA))
  okB <- !any(is.na(invB))
  phi <- ifelse(O > 0, NTB / O, 0)
  phi[!is.finite(phi)] <- 0
  safe_ratio <- function(x) {
    m <- mean(x, na.rm = TRUE)
    if (!is.finite(m) || m == 0) return(rep(NA_real_, length(x)))
    x / m
  }
  if (okA) {
    OM_L  <- colSums(invA)
    FL_L  <- rowSums(invA)
    BL_L  <- colSums(invA)
    FPD_L <- safe_ratio(FL_L)
    BPD_L <- safe_ratio(BL_L)
    VAM_L <- as.numeric(phi %*% invA)
  } else {
    OM_L <- FL_L <- BL_L <- FPD_L <- BPD_L <- VAM_L <- rep(NA_real_, n)
  }
  if (okB) {
    OM_G  <- colSums(invB)
    FL_G  <- rowSums(invB)
    BL_G  <- colSums(invB)
    FPD_G <- safe_ratio(FL_G)
    BPD_G <- safe_ratio(BL_G)
    VAM_G <- as.numeric(phi %*% invB)
  } else {
    OM_G <- FL_G <- BL_G <- FPD_G <- BPD_G <- VAM_G <- rep(NA_real_, n)
  }
  list(
    A = A, B = B, invA = invA, invB = invB, phi = phi,
    OM = OM_L, VAM = VAM_L, FL = FL_L, BL = BL_L, FPD = FPD_L, BPD = BPD_L,
    OM_L = OM_L, VAM_L = VAM_L, FL_L = FL_L, BL_L = BL_L, FPD_L = FPD_L, BPD_L = BPD_L,
    OM_G = OM_G, VAM_G = VAM_G, FL_G = FL_G, BL_G = BL_G, FPD_G = FPD_G, BPD_G = BPD_G,
    okA = okA, okB = okB
  )
}

compute_quadrant2 <- function(tbl) {
  fd <- tbl$fd_new
  Z <- as.matrix(tbl$Z)
  n <- nrow(Z)
  get_fd <- function(key) {
    v <- fd[[key]]
    if (is.null(v) || length(v) == 0) rep(NA_real_, n) else as.numeric(v)
  }
  k3071 <- get_fd("3071")
  k3072 <- get_fd("3072")
  ada_ekspor <- !is.null(fd[["3071"]]) && !is.null(fd[["3072"]])
  k3080 <- if (ada_ekspor) k3071 + k3072 else rep(NA_real_, n)
  list(
    k1800 = rowSums(Z),
    k3011 = get_fd("3011"),
    k3012 = get_fd("3012"),
    k3020 = get_fd("3020"),
    k3030 = get_fd("3030"),
    k3041 = get_fd("3041"),
    k3071 = k3071,
    k3072 = k3072,
    k3080 = k3080,
    k3090 = as.numeric(tbl$konsumsi),
    k3100 = as.numeric(tbl$output),
    k190d = colSums(Z),
    k1900 = colSums(Z) + as.numeric(tbl$impor_ln) + as.numeric(tbl$impor_ap),
    has_fd = !is.null(fd[["3011"]])
  )
}

compute_labor <- function(L, tbl, invA) {
  O <- tbl$output
  prod  <- ifelse(L > 0, O / L, 0)
  lcoef <- ifelse(O > 0, L / O, 0)
  list(
    L = L,
    prod = prod,
    l = lcoef,
    LM = as.numeric(prod %*% invA),
    EM = as.numeric(lcoef %*% invA)
  )
}

build_narratives_model <- function(an, codes, names_sektor, model = c("leontief", "ghosh")) {
  model <- match.arg(model)
  nm <- ifelse(!is.na(names_sektor) & names_sektor != "", names_sektor, codes)
  if (model == "leontief") {
    ok   <- an$okA
    om   <- an$OM_L
    vam  <- an$VAM_L
    fpd  <- an$FPD_L
    bpd  <- an$BPD_L
    label <- "Inverse Leontief"
  } else {
    ok   <- an$okB
    om   <- an$OM_G
    vam  <- an$VAM_G
    fpd  <- an$FPD_G
    bpd  <- an$BPD_G
    label <- "Inverse Ghosh"
  }
  if (!isTRUE(ok)) {
    return(paste0("Matriks ", label, " singular — multiplier & linkage tidak dapat dihitung."))
  }
  out <- c()
  valid <- !is.na(fpd) & !is.na(bpd)
  gt <- nm[valid & fpd > 1 & bpd > 1]
  lt <- nm[valid & fpd < 1 & bpd < 1]
  if (length(gt) > 0) {
    out <- c(out, paste0(
      "[", label, "] Sektor dengan FPD > 1 dan BPD > 1 (kuat menarik hulu & mendorong hilir): ",
      paste(gt, collapse = "; "), "."
    ))
  }
  if (length(lt) > 0) {
    out <- c(out, paste0(
      "[", label, "] Sektor dengan FPD < 1 dan BPD < 1: ",
      paste(lt, collapse = "; "), "."
    ))
  }
  if (length(nm) >= 2) {
    phi <- an$phi
    if (all(is.finite(phi))) {
      ord <- order(phi, decreasing = TRUE)
      out <- c(out, sprintf(
        "[%s] Koefisien nilai tambah terbesar: %s (%.2f%%) dan %s (%.2f%%).",
        label, nm[ord[1]], phi[ord[1]] * 100, nm[ord[2]], phi[ord[2]] * 100
      ))
      ord2 <- order(phi)
      out <- c(out, sprintf(
        "[%s] Koefisien nilai tambah terkecil: %s (%.2f%%) dan %s (%.2f%%).",
        label, nm[ord2[1]], phi[ord2[1]] * 100, nm[ord2[2]], phi[ord2[2]] * 100
      ))
    }
  }
  finite_om <- which(is.finite(om))
  if (length(finite_om) > 0) {
    i_hi <- finite_om[which.max(om[finite_om])]
    i_lo <- finite_om[which.min(om[finite_om])]
    out <- c(out, sprintf(
      "[%s] Output Multiplier terbesar: %s (%.2f) — kenaikan permintaan akhir/primary input Rp1 juta menaikkan output seluruh sektor Rp%.2f juta.",
      label, nm[i_hi], om[i_hi], om[i_hi]
    ))
    out <- c(out, sprintf(
      "[%s] Output Multiplier terkecil: %s (%.2f).",
      label, nm[i_lo], om[i_lo]
    ))
  }
  finite_vam <- which(is.finite(vam))
  if (length(finite_vam) > 0) {
    v_hi <- finite_vam[which.max(vam[finite_vam])]
    v_lo <- finite_vam[which.min(vam[finite_vam])]
    out <- c(out, sprintf(
      "[%s] Value Added Multiplier terbesar: %s (%.3f) — kenaikan permintaan akhir/primary input Rp1 juta menghasilkan nilai tambah Rp%.3f juta.",
      label, nm[v_hi], vam[v_hi], vam[v_hi]
    ))
    out <- c(out, sprintf(
      "[%s] Value Added Multiplier terkecil: %s (%.3f).",
      label, nm[v_lo], vam[v_lo]
    ))
  }
  out
}

build_narratives <- function(an, codes, names_sektor) {
  build_narratives_model(an, codes, names_sektor, model = "leontief")
}

### =============================================================================
### GRAFIK SEBARAN FPD vs BPD (MENYERUPAI REFERENSI EXCEL)
### =============================================================================
make_fpd_bpd_df <- function(fpd, bpd, tbl) {
  df <- data.frame(
    Kode   = template_codes(tbl$codes),
    Sektor = ifelse(!is.na(tbl$names) & tbl$names != "", tbl$names, tbl$codes),
    FPD    = fpd,
    BPD    = bpd,
    stringsAsFactors = FALSE
  )
  df <- df[is.finite(df$FPD) & is.finite(df$BPD), ]
  df
}

build_fpd_bpd_plot <- function(df, judul) {
  xr <- range(c(0.40, 1.80, df$FPD), na.rm = TRUE)
  yr <- range(c(0.60, 1.30, df$BPD), na.rm = TRUE)
  xpad <- 0.05 * (xr[2] - xr[1])
  ypad <- 0.05 * (yr[2] - yr[1])
  xr <- c(xr[1] - xpad, xr[2] + xpad)
  yr <- c(yr[1] - ypad, yr[2] + ypad)
  p <- plot_ly(
    data = df,
    x = ~FPD,
    y = ~BPD,
    type = "scatter",
    mode = "markers+text",
    text = ~Kode,
    textposition = "middle right",
    textfont = list(size = 10, color = "#1F3864"),
    marker = list(color = "#4472C4", size = 7),
    customdata = ~Sektor,
    hovertemplate = paste0(
      "<b>%{text}</b><br>%{customdata}<br>",
      "FPD: %{x:.3f}<br>BPD: %{y:.3f}<extra></extra>"
    )
  )
  p %>%
    layout(
      title = list(
        text = paste0("Sebaran FPD vs BPD (", judul, ")"),
        font = list(size = 15, color = "#0F172A")
      ),
      xaxis = list(
        title = "FPD",
        range = xr,
        dtick = 0.2,
        tickformat = ".2f",
        zeroline = FALSE,
        showgrid = FALSE,
        ticks = "outside",
        tickcolor = "#9CA3AF",
        tickfont = list(size = 10, color = "#334155")
      ),
      yaxis = list(
        title = "BPD",
        range = yr,
        dtick = 0.1,
        tickformat = ".2f",
        zeroline = FALSE,
        showgrid = FALSE,
        ticks = "outside",
        tickcolor = "#9CA3AF",
        tickfont = list(size = 10, color = "#334155")
      ),
      shapes = list(
        list(
          type = "line", x0 = 1, x1 = 1, y0 = yr[1], y1 = yr[2],
          line = list(color = "#8C8C8C", width = 1.2)
        ),
        list(
          type = "line", x0 = xr[1], x1 = xr[2], y0 = 1, y1 = 1,
          line = list(color = "#8C8C8C", width = 1.2)
        )
      ),
      plot_bgcolor = "#FFFFFF",
      paper_bgcolor = "#FFFFFF",
      showlegend = FALSE,
      margin = list(t = 60, b = 55, l = 60, r = 30)
    )
}

### =============================================================================
### HELPER TEMPLATE & STYLE EXCEL
### =============================================================================
template_codes <- function(codes) {
  codes <- trimws(as.character(codes))
  codes[codes == "M,N"] <- "MN"
  codes[codes == "R,S,T,U"] <- "RSTU"
  codes
}

tpl_safe_vector <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  x[!is.finite(x)] <- 0
  x
}

tpl_safe_matrix <- function(M) {
  dm <- dim(M)
  if (is.null(dm)) return(matrix(0, 0, 0))
  M <- matrix(
    suppressWarnings(as.numeric(as.matrix(M))),
    nrow = dm[1],
    ncol = dm[2]
  )
  M[!is.finite(M)] <- 0
  M
}

st_title  <- createStyle(fontSize = 12, textDecoration = "bold", fgFill = "#1E293B", fontColour = "#FFFFFF")
st_hdr    <- createStyle(textDecoration = "bold", fgFill = "#1E293B", fontColour = "#FFFFFF",
                         wrapText = TRUE, halign = "center", valign = "center", border = "TopBottomLeftRight")
st_code   <- createStyle(textDecoration = "bold", fgFill = "#334155", fontColour = "#FFFFFF",
                         halign = "center", border = "TopBottomLeftRight")
st_left   <- createStyle(border = "TopBottomLeftRight")
st_num    <- createStyle(numFmt = "#,##0.00", border = "TopBottomLeftRight")
st_num4   <- createStyle(numFmt = "#,##0.0000", border = "TopBottomLeftRight")
st_coef   <- createStyle(numFmt = "0.0000000000", border = "TopBottomLeftRight")
st_agg    <- createStyle(numFmt = "#,##0.00", textDecoration = "bold", fgFill = "#F1F5F9", border = "TopBottomLeftRight")

### =============================================================================
### GRAFIK STATIS FPD vs BPD UNTUK OUTPUT EXCEL
### =============================================================================
make_fpd_bpd_ggplot <- function(df, judul) {
  xr <- range(c(0.40, 1.80, df$FPD), na.rm = TRUE)
  yr <- range(c(0.60, 1.30, df$BPD), na.rm = TRUE)
  xpad <- 0.05 * diff(xr)
  ypad <- 0.05 * diff(yr)
  xr <- xr + c(-xpad, xpad)
  yr <- yr + c(-ypad, ypad)
  ggplot2::ggplot(df, ggplot2::aes(x = FPD, y = BPD, label = Kode)) +
    ggplot2::geom_vline(xintercept = 1, colour = "#8C8C8C") +
    ggplot2::geom_hline(yintercept = 1, colour = "#8C8C8C") +
    ggplot2::geom_point(colour = "#4472C4", size = 2.4) +
    ggplot2::geom_text(
      hjust = 0,
      vjust = 0.5,
      nudge_x = 0.02,
      size = 3.1,
      colour = "#1F3864"
    ) +
    ggplot2::labs(
      title = paste0("Sebaran FPD vs BPD (", judul, ")"),
      x = "FPD",
      y = "BPD"
    ) +
    ggplot2::coord_cartesian(xlim = xr, ylim = yr) +
    ggplot2::theme_minimal() +
    ggplot2::theme(
      plot.background = ggplot2::element_rect(fill = "white", colour = NA),
      panel.grid = ggplot2::element_blank(),
      axis.line = ggplot2::element_line(colour = "#9CA3AF"),
      axis.ticks = ggplot2::element_line(colour = "#9CA3AF"),
      axis.text = ggplot2::element_text(size = 9, colour = "#334155"),
      axis.title = ggplot2::element_text(size = 10, colour = "#334155"),
      plot.title = ggplot2::element_text(face = "bold", size = 13, colour = "#0F172A"),
      plot.margin = ggplot2::margin(12, 18, 12, 12)
    )
}

save_fpd_bpd_chart <- function(df, judul, prefix = "fpd_bpd",
                               width_cm = 18, height_cm = 12) {
  if (is.null(df) || nrow(df) == 0) return(NULL)
  tryCatch({
    p <- make_fpd_bpd_ggplot(df, judul)
    f <- tempfile(fileext = ".png", pattern = prefix)
    ggplot2::ggsave(
      filename = f,
      plot = p,
      width = width_cm,
      height = height_cm,
      units = "cm",
      dpi = 150
    )
    f
  }, error = function(e) {
    NULL
  })
}

write_template_chart_sheet <- function(wb, sh, title, img,
                                       width_cm = 18, height_cm = 12) {
  addWorksheet(wb, sh)
  writeData(wb, sh, title, startRow = 1, startCol = 1)
  mergeCells(wb, sh, cols = 1:12, rows = 1)
  addStyle(wb, sh, st_title, rows = 1, cols = 1)
  if (!is.null(img) && file.exists(img)) {
    tryCatch(
      {
        insertImage(
          wb,
          sh,
          img,
          startRow = 3,
          startCol = 1,
          width = width_cm,
          height = height_cm,
          units = "cm"
        )
      },
      error = function(e) {
        insertImage(
          wb,
          sh,
          img,
          startRow = 3,
          startCol = 1,
          width = width_cm / 2.54,
          height = height_cm / 2.54
        )
      }
    )
  } else {
    writeData(
      wb,
      sh,
      "Grafik tidak tersedia karena nilai FPD/BPD tidak valid.",
      startRow = 3,
      startCol = 1
    )
  }
  setColWidths(wb, sh, cols = 1:14, widths = rep(11, 14))
}

### =============================================================================
### GENERATE TEMPLATE INPUT
### =============================================================================
generate_template_io <- function(codes, names_sektor) {
  n <- length(codes)
  wb <- createWorkbook()
  addWorksheet(wb, "Tabel IO")
  sh <- "Tabel IO"
  for (i in 1:(n + 1)) writeData(wb, sh, "MATRIKS TABEL IO - TAHUN DASAR", startRow = i, startCol = 1)
  meta_row <- n + 4
  writeData(wb, sh, "Wilayah:", startRow = meta_row, startCol = 1)
  writeData(wb, sh, "Tahun Dasar:", startRow = meta_row, startCol = 3)
  writeData(wb, sh, "Jumlah Sektor:", startRow = meta_row + 1, startCol = 5)
  writeData(wb, sh, n, startRow = meta_row + 1, startCol = 6)
  hdr_row <- meta_row + 3
  writeData(wb, sh, "Kode", startRow = hdr_row, startCol = 1)
  writeData(wb, sh, "Nama Sektor", startRow = hdr_row, startCol = 2)
  for (j in seq_len(n)) writeData(wb, sh, names_sektor[j], startRow = hdr_row, startCol = 2 + j)
  writeData(wb, sh, "Total Konsumsi Akhir", startRow = hdr_row, startCol = 3 + n)
  writeData(wb, sh, "Total Output", startRow = hdr_row, startCol = 4 + n)
  code_row <- hdr_row + 1
  writeData(wb, sh, "Kode", startRow = code_row, startCol = 1)
  writeData(wb, sh, "Nama Sektor", startRow = code_row, startCol = 2)
  for (j in seq_len(n)) writeData(wb, sh, codes[j], startRow = code_row, startCol = 2 + j)
  writeData(wb, sh, 3090, startRow = code_row, startCol = 3 + n)
  writeData(wb, sh, 3100, startRow = code_row, startCol = 4 + n)
  for (i in seq_len(n)) {
    r <- code_row + i
    writeData(wb, sh, codes[i], startRow = r, startCol = 1)
    writeData(wb, sh, names_sektor[i], startRow = r, startCol = 2)
    writeData(wb, sh, matrix(0, 1, n), startRow = r, startCol = 3, colNames = FALSE)
    writeData(wb, sh, 0, startRow = r, startCol = 3 + n)
    writeData(wb, sh, 0, startRow = r, startCol = 4 + n)
  }
  agg_start <- code_row + n + 1
  for (k in 0:3) {
    cd <- c("2000", "2001", "2090", "2100")[k + 1]
    nm <- c(
      "Input Antara Impor Luar Negeri",
      "Input Antara Impor Antar Provinsi",
      "Nilai Tambah Bruto",
      "Total Input"
    )[k + 1]
    writeData(wb, sh, cd, startRow = agg_start + k, startCol = 1)
    writeData(wb, sh, nm, startRow = agg_start + k, startCol = 2)
    writeData(wb, sh, matrix(0, 1, n), startRow = agg_start + k, startCol = 3, colNames = FALSE)
  }
  hs <- createStyle(textDecoration = "bold", fgFill = "#1E293B", fontColour = "#FFFFFF")
  addStyle(wb, sh, hs, rows = hdr_row:code_row, cols = 1:(4 + n), gridExpand = TRUE)
  setColWidths(wb, sh, cols = 1:(4 + n), widths = c(8, 30, rep(12, n), 16, 14))
  wb
}

generate_template_target <- function(codes, names_sektor) {
  n <- length(codes)
  wb <- createWorkbook()
  addWorksheet(wb, "Tabel Target")
  sh <- "Tabel Target"
  for (i in 1:4) writeData(wb, sh, "DATA TAHUN TARGET", startRow = i, startCol = 1)
  writeData(wb, sh, "Wilayah:", startRow = 6, startCol = 1)
  writeData(wb, sh, "Tahun Target:", startRow = 6, startCol = 3)
  hdr_row <- 8
  writeData(wb, sh, "Kode", startRow = hdr_row, startCol = 1)
  writeData(wb, sh, "Nama Sektor", startRow = hdr_row, startCol = 2)
  writeData(wb, sh, "Total Input", startRow = hdr_row, startCol = 3)
  writeData(wb, sh, "Total Output", startRow = hdr_row, startCol = 4)
  for (i in seq_len(n)) {
    r <- hdr_row + i
    writeData(wb, sh, codes[i], startRow = r, startCol = 1)
    writeData(wb, sh, names_sektor[i], startRow = r, startCol = 2)
    writeData(wb, sh, 0, startRow = r, startCol = 3)
    writeData(wb, sh, 0, startRow = r, startCol = 4)
  }
  hs <- createStyle(textDecoration = "bold", fgFill = "#1E293B", fontColour = "#FFFFFF")
  addStyle(wb, sh, hs, rows = hdr_row, cols = 1:4, gridExpand = TRUE)
  setColWidths(wb, sh, cols = 1:4, widths = c(8, 35, 16, 16))
  wb
}

generate_template_labor <- function(codes = CODES_17) {
  wb <- createWorkbook()
  addWorksheet(wb, "Tenaga Kerja")
  sh <- "Tenaga Kerja"
  codes_out <- template_codes(codes)
  writeData(wb, sh, "Kode", startRow = 1, startCol = 1)
  writeData(wb, sh, LABOR_HEADER_TITLE, startRow = 1, startCol = 2)
  for (i in seq_along(codes_out)) {
    r <- 1 + i
    writeData(wb, sh, codes_out[i], startRow = r, startCol = 1)
    writeData(wb, sh, 0, startRow = r, startCol = 2)
  }
  hs <- createStyle(
    textDecoration = "bold",
    fgFill = "#1E293B",
    fontColour = "#FFFFFF",
    wrapText = TRUE,
    valign = "center"
  )
  addStyle(wb, sh, hs, rows = 1, cols = 1:2, gridExpand = TRUE)
  addStyle(wb, sh, createStyle(numFmt = "#,##0"), rows = 2:(1 + length(codes_out)), cols = 2, gridExpand = TRUE)
  setColWidths(wb, sh, cols = 1, widths = 10)
  setColWidths(wb, sh, cols = 2, widths = 60)
  setRowHeights(wb, sh, rows = 1, heights = 30)
  wb
}

### =============================================================================
### OUTPUT EXCEL
### =============================================================================
write_template_update_io_sheet <- function(wb, tbl, q2, M = NULL,
                                           sh = "Update IO",
                                           title = "UPDATING MATRIKS TABEL IO") {
  addWorksheet(wb, sh)
  codes <- template_codes(tbl$codes)
  names_sektor <- ifelse(is.na(tbl$names) | trimws(tbl$names) == "", codes, tbl$names)
  n <- length(codes)
  Z <- tpl_safe_matrix(if (is.null(M)) tbl$Z else M)
  last_col <- 5 + n
  writeData(wb, sh, title, startRow = 1, startCol = 1)
  if (last_col >= 1) mergeCells(wb, sh, cols = 1:last_col, rows = 1)
  addStyle(wb, sh, st_title, rows = 1, cols = 1)
  writeData(wb, sh, "Wilayah:", startRow = 3, startCol = 1)
  writeData(wb, sh, "Tahun Dasar:", startRow = 3, startCol = 3)
  hdr1 <- 5
  hdr2 <- 6
  r0 <- 7
  header_names <- c(
    "Deskripsi", "Kode", names_sektor,
    "Total Permintaan Antara", "Total Konsumsi Akhir", "Total Output"
  )
  header_codes <- c(
    "Deskripsi", "Kode", codes,
    "1800", "3090", "3100"
  )
  writeData(
    wb, sh,
    as.data.frame(t(header_names), stringsAsFactors = FALSE),
    startRow = hdr1, startCol = 1, colNames = FALSE
  )
  writeData(
    wb, sh,
    as.data.frame(t(header_codes), stringsAsFactors = FALSE),
    startRow = hdr2, startCol = 1, colNames = FALSE
  )
  getv <- function(vec, i) {
    if (is.null(vec) || length(vec) < i) return(0)
    tpl_safe_vector(vec[i])
  }
  for (i in seq_len(n)) {
    r <- r0 + i - 1
    writeData(wb, sh, names_sektor[i], startRow = r, startCol = 1)
    writeData(wb, sh, codes[i], startRow = r, startCol = 2)
    writeData(
      wb, sh,
      as.data.frame(t(as.numeric(Z[i, ])), stringsAsFactors = FALSE),
      startRow = r, startCol = 3, colNames = FALSE
    )
    writeData(wb, sh, getv(q2$k1800, i), startRow = r, startCol = 3 + n)
    writeData(wb, sh, getv(q2$k3090, i), startRow = r, startCol = 4 + n)
    writeData(wb, sh, getv(tbl$output, i), startRow = r, startCol = 5 + n)
  }
  agg <- list(
    list(code = "190d", name = "Input Antara Domestik", vals = colSums(Z)),
    list(code = "2000", name = "Input Antara Impor Luar Negeri", vals = tpl_safe_vector(tbl$impor_ln)),
    list(code = "2001", name = "Input Antara Impor Antar Provinsi", vals = tpl_safe_vector(tbl$impor_ap)),
    list(
      code = "1900",
      name = "Total Input Antara",
      vals = colSums(Z) + tpl_safe_vector(tbl$impor_ln) + tpl_safe_vector(tbl$impor_ap)
    ),
    list(code = "2090", name = "Nilai Tambah Bruto", vals = tpl_safe_vector(tbl$ntb)),
    list(code = "2100", name = "Total Input", vals = tpl_safe_vector(tbl$input))
  )
  for (k in seq_along(agg)) {
    r <- r0 + n + k - 1
    writeData(wb, sh, agg[[k]]$name, startRow = r, startCol = 1)
    writeData(wb, sh, agg[[k]]$code, startRow = r, startCol = 2)
    writeData(
      wb, sh,
      as.data.frame(t(as.numeric(agg[[k]]$vals)), stringsAsFactors = FALSE),
      startRow = r, startCol = 3, colNames = FALSE
    )
  }
  addStyle(wb, sh, st_hdr, rows = hdr1, cols = 1:last_col, gridExpand = TRUE)
  addStyle(wb, sh, st_code, rows = hdr2, cols = 1:last_col, gridExpand = TRUE)
  addStyle(wb, sh, st_left, rows = r0:(r0 + n - 1), cols = 1:2, gridExpand = TRUE)
  addStyle(wb, sh, st_num, rows = r0:(r0 + n - 1), cols = 3:last_col, gridExpand = TRUE)
  addStyle(
    wb, sh, st_agg,
    rows = (r0 + n):(r0 + n + length(agg) - 1),
    cols = 1:2, gridExpand = TRUE
  )
  addStyle(
    wb, sh, st_agg,
    rows = (r0 + n):(r0 + n + length(agg) - 1),
    cols = 3:(2 + n), gridExpand = TRUE
  )
  setColWidths(wb, sh, cols = 1, widths = 38)
  setColWidths(wb, sh, cols = 2, widths = 9)
  setColWidths(wb, sh, cols = 3:(2 + n), widths = rep(13, n))
  setColWidths(wb, sh, cols = (3 + n):last_col, widths = rep(20, 3))
  freezePane(wb, sh, firstActiveRow = r0, firstActiveCol = 3)
}

write_template_matrix_sheet <- function(wb, sh, title, M, codes, names_sektor,
                                        add_total_avg = FALSE,
                                        total_value = NULL,
                                        avg_value = NULL) {
  n <- length(codes)
  codes <- template_codes(codes)
  names_sektor <- ifelse(
    is.na(names_sektor) | trimws(names_sektor) == "",
    codes,
    names_sektor
  )
  M <- tpl_safe_matrix(M)
  last_matrix_col <- 2 + n
  last_col <- last_matrix_col + as.integer(add_total_avg)
  addWorksheet(wb, sh)
  writeData(wb, sh, title, startRow = 1, startCol = 1)
  if (last_col >= 1) mergeCells(wb, sh, cols = 1:last_col, rows = 1)
  addStyle(wb, sh, st_title, rows = 1, cols = 1)
  writeData(
    wb, sh,
    as.data.frame(t(c("Deskripsi", "Kode", names_sektor)), stringsAsFactors = FALSE),
    startRow = 2, startCol = 1, colNames = FALSE
  )
  writeData(
    wb, sh,
    as.data.frame(t(c("Deskripsi", "Kode", codes)), stringsAsFactors = FALSE),
    startRow = 3, startCol = 1, colNames = FALSE
  )
  for (i in seq_len(n)) {
    r <- 3 + i
    writeData(wb, sh, names_sektor[i], startRow = r, startCol = 1)
    writeData(wb, sh, codes[i], startRow = r, startCol = 2)
    writeData(
      wb, sh,
      as.data.frame(t(as.numeric(M[i, ])), stringsAsFactors = FALSE),
      startRow = r, startCol = 3, colNames = FALSE
    )
  }
  addStyle(wb, sh, st_hdr, rows = 2, cols = 1:last_matrix_col, gridExpand = TRUE)
  addStyle(wb, sh, st_code, rows = 3, cols = 1:last_matrix_col, gridExpand = TRUE)
  addStyle(wb, sh, st_left, rows = 4:(3 + n), cols = 1:2, gridExpand = TRUE)
  addStyle(wb, sh, st_coef, rows = 4:(3 + n), cols = 3:last_matrix_col, gridExpand = TRUE)
  if (add_total_avg) {
    r1 <- 4 + n
    if (is.null(total_value)) total_value <- sum(M, na.rm = TRUE)
    if (is.null(avg_value)) avg_value <- mean(colSums(M, na.rm = TRUE), na.rm = TRUE)
    total_value <- tpl_safe_vector(total_value)
    avg_value <- tpl_safe_vector(avg_value)
    writeData(wb, sh, "Total Semua Sektor", startRow = r1, startCol = 2)
    writeData(wb, sh, total_value, startRow = r1, startCol = 3 + n)
    writeData(wb, sh, "Rata-rata", startRow = r1 + 1, startCol = 2)
    writeData(wb, sh, avg_value, startRow = r1 + 1, startCol = 3 + n)
    addStyle(wb, sh, st_left, rows = r1:(r1 + 1), cols = 2, gridExpand = TRUE)
    addStyle(wb, sh, st_num4, rows = r1:(r1 + 1), cols = 3 + n, gridExpand = TRUE)
  }
  setColWidths(wb, sh, cols = 1, widths = 38)
  setColWidths(wb, sh, cols = 2, widths = 9)
  setColWidths(wb, sh, cols = 3:last_matrix_col, widths = rep(13, n))
  if (add_total_avg) {
    setColWidths(wb, sh, cols = 3 + n, widths = 18)
  }
  freezePane(wb, sh, firstActiveRow = 4, firstActiveCol = 3)
}

write_template_simple_table <- function(wb, sh, title, header, df, num_cols = NULL) {
  addWorksheet(wb, sh)
  writeData(wb, sh, title, startRow = 1, startCol = 1)
  if (length(header) >= 1) mergeCells(wb, sh, cols = seq_along(header), rows = 1)
  addStyle(wb, sh, st_title, rows = 1, cols = 1)
  writeData(
    wb, sh,
    as.data.frame(t(header), stringsAsFactors = FALSE),
    startRow = 2, startCol = 1, colNames = FALSE
  )
  addStyle(wb, sh, st_hdr, rows = 2, cols = seq_along(header), gridExpand = TRUE)
  if (nrow(df) > 0) {
    writeData(wb, sh, df, startRow = 3, startCol = 1, colNames = FALSE)
    addStyle(
      wb, sh, st_left,
      rows = 3:(2 + nrow(df)),
      cols = seq_along(header),
      gridExpand = TRUE
    )
    if (!is.null(num_cols) && length(num_cols) > 0) {
      addStyle(
        wb, sh, st_num4,
        rows = 3:(2 + nrow(df)),
        cols = num_cols,
        gridExpand = TRUE
      )
    }
  }
  setColWidths(
    wb, sh,
    cols = seq_along(header),
    widths = c(10, rep(28, length(header) - 1))
  )
}

write_template_labor_sheet <- function(wb, tbl, labor) {
  sh <- "Labor Effect"
  addWorksheet(wb, sh)
  codes <- template_codes(tbl$codes)
  n <- length(codes)
  writeData(wb, sh, "Labor Effect", startRow = 1, startCol = 1)
  mergeCells(wb, sh, cols = 1:6, rows = 1)
  addStyle(wb, sh, st_title, rows = 1, cols = 1)
  header <- c(
    "Kode",
    LABOR_HEADER_TITLE,
    "Produktivitas, Tenaga Kerja/Output",
    "",
    "Sektor",
    "Labor Multiplier"
  )
  writeData(
    wb, sh,
    as.data.frame(t(header), stringsAsFactors = FALSE),
    startRow = 2, startCol = 1, colNames = FALSE
  )
  addStyle(wb, sh, st_hdr, rows = 2, cols = 1:6, gridExpand = TRUE)
  for (i in seq_len(n)) {
    r <- 2 + i
    writeData(wb, sh, codes[i], startRow = r, startCol = 1)
    if (!is.null(labor) && length(labor$L) >= i && is.finite(labor$L[i])) {
      writeData(wb, sh, labor$L[i], startRow = r, startCol = 2)
    } else {
      writeData(wb, sh, "", startRow = r, startCol = 2)
    }
    if (!is.null(labor) && length(labor$prod) >= i && is.finite(labor$prod[i])) {
      writeData(wb, sh, labor$prod[i], startRow = r, startCol = 3)
    } else {
      writeData(wb, sh, "", startRow = r, startCol = 3)
    }
    writeData(wb, sh, "", startRow = r, startCol = 4)
    writeData(wb, sh, codes[i], startRow = r, startCol = 5)
    if (!is.null(labor) && length(labor$LM) >= i && is.finite(labor$LM[i])) {
      writeData(wb, sh, labor$LM[i], startRow = r, startCol = 6)
    } else {
      writeData(wb, sh, "", startRow = r, startCol = 6)
    }
  }
  addStyle(wb, sh, st_left, rows = 3:(2 + n), cols = 1:6, gridExpand = TRUE)
  addStyle(wb, sh, st_num, rows = 3:(2 + n), cols = c(2, 3, 6), gridExpand = TRUE)
  setColWidths(wb, sh, cols = 1, widths = 8)
  setColWidths(wb, sh, cols = 2, widths = 55)
  setColWidths(wb, sh, cols = 3, widths = 25)
  setColWidths(wb, sh, cols = 4, widths = 3)
  setColWidths(wb, sh, cols = 5, widths = 10)
  setColWidths(wb, sh, cols = 6, widths = 18)
}

write_template_pmtb_sheet <- function(wb, tbl, q2) {
  sh <- "Rasio Per Output"
  addWorksheet(wb, sh)
  codes <- template_codes(tbl$codes)
  n <- length(codes)
  writeData(wb, sh, "Rasio Per Output", startRow = 1, startCol = 1)
  mergeCells(wb, sh, cols = 1:2, rows = 1)
  addStyle(wb, sh, st_title, rows = 1, cols = 1)
  header <- c("Kode", "Ratio PMTB/Output")
  writeData(
    wb, sh,
    as.data.frame(t(header), stringsAsFactors = FALSE),
    startRow = 2, startCol = 1, colNames = FALSE
  )
  addStyle(wb, sh, st_hdr, rows = 2, cols = 1:2, gridExpand = TRUE)
  for (i in seq_len(n)) {
    r <- 2 + i
    writeData(wb, sh, codes[i], startRow = r, startCol = 1)
    val <- NA_real_
    if (!is.null(q2$k3030) && length(q2$k3030) >= i) {
      pmtb <- suppressWarnings(as.numeric(q2$k3030[i]))
      out <- suppressWarnings(as.numeric(tbl$output[i]))
      if (is.finite(pmtb) && is.finite(out) && out != 0) {
        val <- pmtb / out
      }
    }
    if (is.finite(val)) {
      writeData(wb, sh, val, startRow = r, startCol = 2)
    } else {
      writeData(wb, sh, "", startRow = r, startCol = 2)
    }
  }
  addStyle(wb, sh, st_left, rows = 3:(2 + n), cols = 1:2, gridExpand = TRUE)
  addStyle(wb, sh, st_num4, rows = 3:(2 + n), cols = 2, gridExpand = TRUE)
  setColWidths(wb, sh, cols = 1, widths = 10)
  setColWidths(wb, sh, cols = 2, widths = 25)
}

build_update_io_only_workbook <- function(h, tbl, q2) {
  wb <- createWorkbook()
  write_template_update_io_sheet(wb, tbl, q2)
  wb
}

build_template_output_workbook <- function(h, tbl, an, q2, labor, images = NULL) {
  wb <- createWorkbook()
  n <- length(tbl$codes)
  write_template_update_io_sheet(wb, tbl, q2)
  df_koef <- data.frame(
    Sektor = template_codes(tbl$codes),
    Koefisien = tpl_safe_vector(an$phi),
    stringsAsFactors = FALSE
  )
  write_template_simple_table(
    wb, "Koefisien Nilai Tambah", "Koefisien Nilai Tambah",
    c("Sektor", "Koefisien Nilai Tambah"), df_koef, num_cols = 2
  )
  I_mat <- diag(n)
  A <- tpl_safe_matrix(an$A)
  R_mat <- tpl_safe_matrix(an$B)
  invA <- tpl_safe_matrix(an$invA)
  invR <- tpl_safe_matrix(an$invB)
  write_template_update_io_sheet(wb, tbl, q2, M = I_mat, sh = "I", title = "I")
  write_template_matrix_sheet(wb, "A", "A", A, tbl$codes, tbl$names)
  write_template_matrix_sheet(wb, "I-A", "I-A", I_mat - A, tbl$codes, tbl$names)
  total_ia <- sum(invA, na.rm = TRUE)
  avg_ia <- if (n > 0) mean(colSums(invA), na.rm = TRUE) else 0
  if (!is.finite(avg_ia)) avg_ia <- 0
  write_template_matrix_sheet(
    wb, "(I-A)^(-1)", "(I-A)^(-1)", invA, tbl$codes, tbl$names,
    add_total_avg = TRUE, total_value = total_ia, avg_value = avg_ia
  )
  write_template_matrix_sheet(wb, "R", "R", R_mat, tbl$codes, tbl$names)
  write_template_matrix_sheet(wb, "(I-R)", "(I-R)", I_mat - R_mat, tbl$codes, tbl$names)
  write_template_matrix_sheet(wb, "(I-R)^(-1)", "(I-R)^(-1)", invR, tbl$codes, tbl$names)
  safe_ratio_excel <- function(x) {
    x <- tpl_safe_vector(x)
    m <- mean(x, na.rm = TRUE)
    if (!is.finite(m) || m == 0) return(rep(0, length(x)))
    x / m
  }
  FL_L <- tpl_safe_vector(rowSums(invA))
  BL_L <- tpl_safe_vector(colSums(invA))
  FPD_L <- safe_ratio_excel(FL_L)
  BPD_L <- safe_ratio_excel(BL_L)
  OM_L <- BL_L
  VAM_L <- tpl_safe_vector(an$VAM_L)
  FL_G <- tpl_safe_vector(rowSums(invR))
  BL_G <- tpl_safe_vector(colSums(invR))
  FPD_G <- safe_ratio_excel(FL_G)
  BPD_G <- safe_ratio_excel(BL_G)
  OM_G <- BL_G
  VAM_G <- tpl_safe_vector(an$VAM_G)
  df_forward <- data.frame(
    Kode = template_codes(tbl$codes),
    FPD = FPD_L,
    FL = FL_L,
    stringsAsFactors = FALSE
  )
  write_template_simple_table(
    wb, "Forward", "Forward (Leontief)",
    c("Kode", "Forward Power of Dispersion", "Forward Linkage"),
    df_forward, num_cols = 2:3
  )
  df_backward <- data.frame(
    Kode = template_codes(tbl$codes),
    BPD = BPD_L,
    BL = BL_L,
    stringsAsFactors = FALSE
  )
  write_template_simple_table(
    wb, "Backward", "Backward (Leontief)",
    c("Kode", "Backward Power of Dispersion", "Backward Linkage"),
    df_backward, num_cols = 2:3
  )
  df_om <- data.frame(
    Kode = template_codes(tbl$codes),
    Output_Multiplier = OM_L,
    stringsAsFactors = FALSE
  )
  write_template_simple_table(
    wb, "Output Multiplier", "Output Multiplier (Leontief)",
    c("Kode", "Output Multiplier"),
    df_om, num_cols = 2
  )
  df_vam <- data.frame(
    Kode = template_codes(tbl$codes),
    Ratio_NTB_Output = tpl_safe_vector(an$phi),
    Nilai_Tambah_Multiplier = VAM_L,
    stringsAsFactors = FALSE
  )
  write_template_simple_table(
    wb, "Nilai Tambah Multiplier", "Nilai Tambah Multiplier (Leontief)",
    c("Kode", "Ratio NTB/Output", "Nilai Tambah Multiplier"),
    df_vam, num_cols = 2:3
  )
  df_forward_g <- data.frame(
    Kode = template_codes(tbl$codes),
    FPD = FPD_G,
    FL = FL_G,
    stringsAsFactors = FALSE
  )
  write_template_simple_table(
    wb, "Forward Ghosh", "Forward (Ghosh)",
    c("Kode", "Forward Power of Dispersion", "Forward Linkage"),
    df_forward_g, num_cols = 2:3
  )
  df_backward_g <- data.frame(
    Kode = template_codes(tbl$codes),
    BPD = BPD_G,
    BL = BL_G,
    stringsAsFactors = FALSE
  )
  write_template_simple_table(
    wb, "Backward Ghosh", "Backward (Ghosh)",
    c("Kode", "Backward Power of Dispersion", "Backward Linkage"),
    df_backward_g, num_cols = 2:3
  )
  df_om_g <- data.frame(
    Kode = template_codes(tbl$codes),
    Output_Multiplier = OM_G,
    stringsAsFactors = FALSE
  )
  write_template_simple_table(
    wb, "Output Mult Ghosh", "Output Multiplier (Ghosh)",
    c("Kode", "Output Multiplier"),
    df_om_g, num_cols = 2
  )
  df_vam_g <- data.frame(
    Kode = template_codes(tbl$codes),
    Ratio_NTB_Output = tpl_safe_vector(an$phi),
    Nilai_Tambah_Multiplier = VAM_G,
    stringsAsFactors = FALSE
  )
  write_template_simple_table(
    wb, "VAM Ghosh", "Nilai Tambah Multiplier (Ghosh)",
    c("Kode", "Ratio NTB/Output", "Nilai Tambah Multiplier"),
    df_vam_g, num_cols = 2:3
  )
  write_template_labor_sheet(wb, tbl, labor)
  write_template_pmtb_sheet(wb, tbl, q2)
  if (!is.null(images)) {
    write_template_chart_sheet(
      wb,
      "Grafik FPD BPD Leontief",
      "Grafik Sebaran FPD vs BPD [Berdasarkan Inverse Leontief]",
      images$leontief,
      width_cm = 18,
      height_cm = 12
    )
    write_template_chart_sheet(
      wb,
      "Grafik FPD BPD Ghosh",
      "Grafik Sebaran FPD vs BPD [Berdasarkan Inverse Ghosh]",
      images$ghosh,
      width_cm = 18,
      height_cm = 12
    )
  }
  wb
}

### =============================================================================
### UI — CSS & JS
### =============================================================================
custom_css <- "
* { box-sizing: border-box; }
html, body { max-width: 100%; }
body {
  margin: 0;
  font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
  background-color: #F1F5F9;
  overflow-x: hidden;
}
.app-wrapper { display: flex; min-height: 100vh; width: 100%; }
.sidebar {
  width: 270px;
  min-width: 270px;
  background: linear-gradient(180deg, #0F172A 0%, #1E293B 100%);
  color: #FFFFFF;
  display: flex;
  flex-direction: column;
  position: fixed;
  top: 0; left: 0; bottom: 0;
  z-index: 100;
  box-shadow: 4px 0 20px rgba(0,0,0,0.15);
}
.sidebar-header { padding: 1.75rem 1.5rem; border-bottom: 1px solid rgba(255,255,255,0.08); }
.sidebar-header h4 { margin: 0 0 0.25rem 0; font-weight: 700; font-size: 1.1rem; color: #FFFFFF; }
.sidebar-header .subtitle { font-size: 0.75rem; color: #94A3B8; text-transform: uppercase; letter-spacing: 0.08em; }
.sidebar-nav { flex: 1; padding: 1rem 0.75rem; overflow-y: auto; }
.nav-section-label {
  font-size: 0.65rem;
  text-transform: uppercase;
  letter-spacing: 0.1em;
  color: #64748B;
  padding: 0.75rem 0.75rem 0.5rem;
  font-weight: 600;
}
.sidebar-item {
  display: flex;
  align-items: center;
  gap: 0.75rem;
  padding: 0.8rem 1rem;
  margin-bottom: 0.25rem;
  border-radius: 0.6rem;
  cursor: pointer;
  transition: all 0.2s ease;
  color: #CBD5E1;
  font-size: 0.9rem;
  font-weight: 500;
  border: none;
  background: transparent;
  width: 100%;
  text-align: left;
}
.sidebar-item:hover { background: rgba(255,255,255,0.06); color: #FFFFFF; }
.sidebar-item.active {
  background: linear-gradient(135deg, #10B981, #059669);
  color: #FFFFFF;
  box-shadow: 0 4px 12px rgba(16,185,129,0.3);
}
.sidebar-item i { width: 20px; text-align: center; font-size: 1rem; }
.sidebar-item[disabled] {
  opacity: 0.45;
  cursor: not-allowed;
}
.main-content {
  flex: 1 1 auto;
  margin-left: 270px;
  padding: 2rem;
  min-height: 100vh;
  width: calc(100% - 270px);
  max-width: calc(100% - 270px);
  min-width: 0;
  overflow-x: hidden;
}
.main-panel { display: none; min-width: 0; max-width: 100%; }
.main-panel.active-panel { display: block; animation: fadeIn 0.3s ease; }
@keyframes fadeIn {
  from { opacity: 0; transform: translateY(8px); }
  to { opacity: 1; transform: translateY(0); }
}
.page-header { margin-bottom: 1.75rem; }
.page-header h3 { font-weight: 700; color: #0F172A; margin-bottom: 0.35rem; font-size: 1.4rem; }
.page-header p { color: #64748B; margin: 0; font-size: 0.9rem; }
.content-card {
  background: #FFFFFF;
  border: 1px solid #E2E8F0;
  border-radius: 0.85rem;
  box-shadow: 0 1px 3px rgba(15,23,42,0.06);
  margin-bottom: 1.5rem;
  overflow: hidden;
  max-width: 100%;
  min-width: 0;
}
.content-card .card-title {
  padding: 1rem 1.5rem;
  font-weight: 600;
  font-size: 0.95rem;
  color: #0F172A;
  border-bottom: 1px solid #F1F5F9;
  display: flex;
  align-items: center;
  gap: 0.6rem;
}
.content-card .card-title i { color: #10B981; }
.content-card .card-body-inner { padding: 1.5rem; max-width: 100%; min-width: 0; overflow-x: auto; }
.title-control { margin-left: auto; display: flex; align-items: center; gap: 0.5rem; }
.title-control span { font-size: 0.75rem; color: #64748B; font-weight: 500; }
.title-control .shiny-input-container { margin-bottom: 0; }
.title-control select {
  padding: 0.3rem 0.5rem;
  font-size: 0.8rem;
  border-radius: 0.4rem;
  border: 1px solid #E2E8F0;
  background: #F8FAFC;
  color: #0F172A;
}
.dataTables_wrapper { width: 100% !important; max-width: 100%; }
.btn-process {
  background: linear-gradient(135deg, #10B981, #059669);
  border: none;
  color: #FFFFFF;
  font-weight: 600;
  font-size: 1rem;
  padding: 0.85rem 2.5rem;
  border-radius: 0.65rem;
  cursor: pointer;
  transition: all 0.2s ease;
  box-shadow: 0 4px 14px rgba(16,185,129,0.35);
  display: inline-flex;
  align-items: center;
  gap: 0.6rem;
}
.btn-process:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(16,185,129,0.45); }
.btn-process[disabled] {
  opacity: 0.55;
  cursor: not-allowed;
  box-shadow: none;
  transform: none;
}
.btn-template {
  background: #FFFFFF;
  border: 2px solid #E2E8F0;
  color: #334155;
  font-weight: 600;
  font-size: 0.88rem;
  padding: 0.7rem 1.5rem;
  border-radius: 0.6rem;
  cursor: pointer;
  transition: all 0.2s ease;
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
}
.btn-template:hover { border-color: #10B981; color: #10B981; background: #F0FDF9; }
.btn-download-result {
  background: linear-gradient(135deg, #1E293B, #0F172A);
  border: none;
  color: #FFFFFF;
  font-weight: 600;
  font-size: 0.95rem;
  padding: 0.8rem 2rem;
  border-radius: 0.65rem;
  cursor: pointer;
  transition: all 0.2s ease;
  display: inline-flex;
  align-items: center;
  gap: 0.6rem;
}
.btn-download-result:hover { transform: translateY(-1px); box-shadow: 0 4px 16px rgba(15,23,42,0.3); }
.status-box {
  padding: 1rem 1.25rem;
  border-radius: 0.65rem;
  font-size: 0.88rem;
  font-weight: 500;
  margin-top: 1.25rem;
  display: flex;
  align-items: flex-start;
  gap: 0.6rem;
  line-height: 1.5;
}
.status-box.status-info { background: #ECFDF5; border: 1px solid #A7F3D0; color: #065F46; }
.status-box.status-warning { background: #FFF7ED; border: 1px solid #FED7AA; color: #9A3412; }
.status-box.status-error { background: #FEF2F2; border: 1px solid #FECACA; color: #991B1B; }
.narasi-list { margin: 0; padding-left: 1.2rem; color: #334155; font-size: 0.88rem; line-height: 1.9; }
.loading-overlay {
  position: fixed;
  top: 0; left: 0; right: 0; bottom: 0;
  background: rgba(15, 23, 42, 0.92);
  backdrop-filter: blur(4px);
  z-index: 99999;
  display: none;
  justify-content: center;
  align-items: center;
  flex-direction: column;
}
.loading-overlay.show { display: flex; }
.loading-card {
  background: #FFFFFF;
  border-radius: 1.25rem;
  padding: 3rem 3.5rem;
  text-align: center;
  box-shadow: 0 25px 60px rgba(0,0,0,0.3);
  max-width: 420px;
  width: 90%;
}
.loading-spinner {
  width: 56px;
  height: 56px;
  border: 4px solid #E2E8F0;
  border-top: 4px solid #10B981;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
  margin: 0 auto 1.5rem;
}
@keyframes spin { to { transform: rotate(360deg); } }
.loading-percent {
  font-size: 2.75rem;
  font-weight: 800;
  color: #0F172A;
  margin-bottom: 0.25rem;
  font-variant-numeric: tabular-nums;
}
.loading-message { font-size: 0.88rem; color: #64748B; margin-bottom: 1.5rem; min-height: 1.2em; }
.progress-track { width: 100%; height: 10px; background: #E2E8F0; border-radius: 99px; overflow: hidden; }
.progress-fill {
  height: 100%;
  width: 0%;
  background: linear-gradient(90deg, #10B981, #059669);
  border-radius: 99px;
  transition: width 0.15s ease;
}
.info-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
  gap: 1rem;
  width: 100%;
  max-width: 100%;
  min-width: 0;
}
.info-item { background: #F8FAFC; border-radius: 0.6rem; padding: 1rem; text-align: center; min-width: 0; }
.info-item .label {
  font-size: 0.72rem;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  color: #64748B;
  margin-bottom: 0.35rem;
}
.info-item .value { font-size: 1.3rem; font-weight: 700; color: #0F172A; word-break: break-word; }
.template-card {
  border: 1px solid #E2E8F0;
  border-radius: 0.75rem;
  padding: 1.25rem;
  display: flex;
  align-items: center;
  gap: 1rem;
  transition: all 0.2s ease;
  background: #FFFFFF;
}
.template-card:hover { border-color: #10B981; box-shadow: 0 4px 12px rgba(16,185,129,0.1); }
.template-icon {
  width: 44px;
  height: 44px;
  background: linear-gradient(135deg, #ECFDF5, #D1FAE5);
  border-radius: 0.65rem;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.2rem;
  color: #059669;
  flex-shrink: 0;
}
.template-info { flex: 1; min-width: 0; }
.template-info h6 { margin: 0 0 0.2rem 0; font-weight: 600; font-size: 0.9rem; color: #0F172A; }
.template-info p { margin: 0; font-size: 0.75rem; color: #64748B; }
#file_base, #file_target, #file_analysis { width: 100%; }
.shiny-input-container:has(#file_base),
.shiny-input-container:has(#file_target),
.shiny-input-container:has(#file_analysis) {
  width: 100% !important;
  max-width: 100% !important;
  min-width: 0 !important;
}
.shiny-input-container:has(#file_base) .input-group,
.shiny-input-container:has(#file_target) .input-group,
.shiny-input-container:has(#file_analysis) .input-group {
  display: flex !important;
  width: 100% !important;
  max-width: 100% !important;
  min-width: 0 !important;
  flex-wrap: nowrap !important;
}
.shiny-input-container:has(#file_base) .input-group-btn,
.shiny-input-container:has(#file_target) .input-group-btn,
.shiny-input-container:has(#file_analysis) .input-group-btn {
  flex: 0 0 88px !important;
  width: 88px !important;
  min-width: 88px !important;
}
.shiny-input-container:has(#file_base) .input-group-btn > .btn,
.shiny-input-container:has(#file_target) .input-group-btn > .btn,
.shiny-input-container:has(#file_analysis) .input-group-btn > .btn {
  width: 100% !important;
  height: 44px !important;
  min-height: 44px !important;
  padding: 0 0.75rem !important;
  white-space: nowrap !important;
}
.shiny-input-container:has(#file_base) .input-group > .form-control,
.shiny-input-container:has(#file_target) .input-group > .form-control,
.shiny-input-container:has(#file_analysis) .input-group > .form-control {
  flex: 1 1 auto !important;
  width: 1% !important;
  max-width: none !important;
  min-width: 0 !important;
  height: 44px !important;
  min-height: 44px !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  white-space: nowrap !important;
}
@media (max-width: 768px) {
  .sidebar { transform: translateX(-100%); }
  .sidebar.open { transform: translateX(0); }
  .main-content { margin-left: 0; width: 100%; max-width: 100%; }
}
"

custom_js <- "
Shiny.addCustomMessageHandler('updateProgress', function(data) {
  var pct = Math.round(data.percent);
  document.getElementById('progress-fill').style.width = pct + '%';
  document.getElementById('progress-percent').textContent = pct + '%';
  document.getElementById('progress-message').textContent = data.message;
});
Shiny.addCustomMessageHandler('showOverlay', function(data) {
  document.getElementById('loading-overlay').classList.add('show');
});
Shiny.addCustomMessageHandler('hideOverlay', function(data) {
  document.getElementById('loading-overlay').classList.remove('show');
});
Shiny.addCustomMessageHandler('switchPanel', function(panelId) {
  switchPanel(panelId);
});
Shiny.addCustomMessageHandler('setHasilAccess', function(enabled) {
  var btn = document.getElementById('nav-panel-hasil');
  if (btn) {
    btn.disabled = !enabled;
  }
  if (!enabled) {
    var active = document.querySelector('.main-panel.active-panel');
    if (active && active.id === 'panel-hasil') {
      switchPanel('panel-analisis');
    }
  }
});
function switchPanel(panelId) {
  var panels = document.querySelectorAll('.main-panel');
  for (var i = 0; i < panels.length; i++) {
    panels[i].classList.remove('active-panel');
  }
  document.getElementById(panelId).classList.add('active-panel');
  var items = document.querySelectorAll('.sidebar-item');
  for (var i = 0; i < items.length; i++) {
    items[i].classList.remove('active');
  }
  document.getElementById('nav-' + panelId).classList.add('active');
}
"

### =============================================================================
### UI
### =============================================================================
ui <- page_fluid(
  theme = bs_theme(
    version = 5,
    bg = "#F1F5F9",
    fg = "#0F172A",
    primary = "#10B981",
    secondary = "#1E293B",
    "font-sans-serif" = "system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif"
  ),
  tags$head(
    tags$style(HTML(custom_css)),
    tags$script(HTML(custom_js))
  ),
  div(
    id = "loading-overlay",
    class = "loading-overlay",
    div(
      class = "loading-card",
      div(class = "loading-spinner"),
      div(id = "progress-percent", class = "loading-percent", "0%"),
      div(id = "progress-message", class = "loading-message", "Memulai proses..."),
      div(
        class = "progress-track",
        div(id = "progress-fill", class = "progress-fill")
      )
    )
  ),
  div(
    class = "app-wrapper",
    div(
      class = "sidebar",
      div(
        class = "sidebar-header",
        h4(icon("table"), " I-O Table Updater"),
        div(class = "subtitle", "Iterative Proportional Fitting")
      ),
      div(
        class = "sidebar-nav",
        div(class = "nav-section-label", "MENU UTAMA"),
        tags$button(
          id = "nav-panel-input",
          class = "sidebar-item active",
          onclick = "switchPanel('panel-input')",
          icon("upload"), " Input Data"
        ),
        tags$button(
          id = "nav-panel-hasil",
          class = "sidebar-item",
          onclick = "if (!this.disabled) switchPanel('panel-hasil');",
          icon("th"), " Hasil Tabel I-O" # IKON-DIPERBARUI
        ),
        tags$button(
          id = "nav-panel-analisis",
          class = "sidebar-item",
          onclick = "switchPanel('panel-analisis')",
          icon("calculator"), " Analisis I-O"
        ),
        tags$button(
          id = "nav-panel-output",
          class = "sidebar-item",
          onclick = "switchPanel('panel-output')",
          icon("download"), " Output"
        ),
        div(class = "nav-section-label", "LAINNYA"),
        tags$button(
          id = "nav-panel-template",
          class = "sidebar-item",
          onclick = "switchPanel('panel-template')",
          icon("copy"), " Template" # IKON-DIPERBARUI
        )
      )
    ),
    div(
      class = "main-content",
      div(
        id = "panel-input",
        class = "main-panel active-panel",
        div(
          class = "page-header",
          h3("Input Data")
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("bolt"), " Jalur 1 — Analisis IO (Tanpa Updating IO)"), # IKON-DIPERBARUI
          div(
            class = "card-body-inner",
            fileInput(
              "file_analysis",
              "Upload Data untuk Analisis",
              accept = c(".xlsx", ".xls"),
              buttonLabel = "Pilih File...",
              placeholder = "Belum ada berkas dipilih"
            )
          )
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("balance-scale"), " Jalur 2 — Updating IO (IPFP)"), # IKON-DIPERBARUI
          div(
            class = "card-body-inner",
            uiOutput("update_validation_ui"),
            div(
              style = "display:grid; grid-template-columns: 1fr 1fr; gap:1.5rem;",
              div(
                fileInput(
                  "file_base",
                  "Upload Tabel IO Tahun Dasar",
                  accept = c(".xlsx", ".xls"),
                  buttonLabel = "Pilih File...",
                  placeholder = "Belum ada berkas dipilih"
                )
              ),
              div(
                fileInput(
                  "file_target",
                  "Upload Data Tahun Target",
                  accept = c(".xlsx", ".xls"),
                  buttonLabel = "Pilih File...",
                  placeholder = "Belum ada berkas dipilih"
                )
              )
            ),
            div(
              style = "display:grid; grid-template-columns: 1fr 1fr; gap:1.5rem;",
              numericInput(
                "max_iter",
                "Maksimum Iterasi",
                value = 1000,
                min = 1,
                max = 100000,
                step = 1
              ),
              numericInput(
                "tol_value",
                "Toleransi Konvergensi",
                value = 0.000001,
                min = 0,
                max = 1,
                step = 0.000001
              )
            ),
            div(
              style = "text-align:center; margin: 0.25rem 0 0.5rem 0;",
              uiOutput("btn_update_ui")
            )
          )
        ),
        uiOutput("sumber_data_ui"),
        uiOutput("versi_ui"),
        uiOutput("status_ui")
      ),
      div(
        id = "panel-hasil",
        class = "main-panel",
        div(class = "page-header", h3("Hasil Tabel I-O")),
        uiOutput("hasil_info_ui"),
        div(
          class = "content-card",
          div(
            class = "card-title",
            icon("clipboard-list"), " Ringkasan Sektor", # IKON-DIPERBARUI
            div(
              class = "title-control",
              tags$span("Tampilkan:"),
              selectInput(
                "rows_ringkasan",
                NULL,
                choices = list(
                  "10" = "10",
                  "25" = "25",
                  "50" = "50",
                  "100" = "100",
                  "Semua" = "-1"
                ),
                selected = "25",
                width = "90px"
              )
            )
          ),
          div(class = "card-body-inner", DTOutput("tbl_ringkasan"))
        ),
        div(
          class = "content-card",
          div(
            class = "card-title",
            icon("border-all"), " Matriks Transaksi Antara (Z)", # IKON-DIPERBARUI
            div(
              class = "title-control",
              tags$span("Tampilkan:"),
              selectInput(
                "rows_matriks",
                NULL,
                choices = list(
                  "10" = "10",
                  "25" = "25",
                  "50" = "50",
                  "100" = "100",
                  "Semua" = "-1"
                ),
                selected = "10",
                width = "90px"
              )
            )
          ),
          div(class = "card-body-inner", DTOutput("tbl_matriks"))
        )
      ),
      div(
        id = "panel-analisis",
        class = "main-panel",
        div(
          class = "page-header",
          h3("Analisis I-O")
        ),
        uiOutput("analisis_status_ui"),
        div(
          class = "content-card",
          div(class = "card-title", icon("square-root-alt"), " Matriks Koefisien & Invers"), # IKON-DIPERBARUI
          div(
            class = "card-body-inner",
            tabsetPanel(
              tabPanel("Koefisien Input (A)", br(), DTOutput("tbl_A")),
              tabPanel("Inverse Leontief (I-A)^-1", br(), DTOutput("tbl_invA")),
              tabPanel("Koefisien Output (R)", br(), DTOutput("tbl_B")),
              tabPanel("Inverse Ghosh (I-R)^-1", br(), DTOutput("tbl_invB"))
            )
          )
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("percent"), " Koefisien Nilai Tambah & Multiplier [Berdasarkan Inverse Leontief]"),
          div(class = "card-body-inner", DTOutput("tbl_multiplier_leontief"))
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("project-diagram"), " Forward-Backward Linkage & Power of Dispersion [Berdasarkan Inverse Leontief]"), # IKON-DIPERBARUI
          div(class = "card-body-inner", DTOutput("tbl_linkage_leontief"))
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("chart-line"), " Grafik Sebaran FPD vs BPD [Berdasarkan Inverse Leontief]"),
          div(class = "card-body-inner", plotlyOutput("grafik_fpd_bpd_leontief", height = "520px"))
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("lightbulb"), " Interpretasi Otomatis [Berdasarkan Inverse Leontief]"),
          div(class = "card-body-inner", uiOutput("narasi_ui_leontief"))
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("percent"), " Koefisien Nilai Tambah & Multiplier [Berdasarkan Inverse Ghosh]"),
          div(class = "card-body-inner", DTOutput("tbl_multiplier_ghosh"))
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("project-diagram"), " Forward-Backward Linkage & Power of Dispersion [Berdasarkan Inverse Ghosh]"), # IKON-DIPERBARUI
          div(class = "card-body-inner", DTOutput("tbl_linkage_ghosh"))
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("chart-line"), " Grafik Sebaran FPD vs BPD [Berdasarkan Inverse Ghosh]"),
          div(class = "card-body-inner", plotlyOutput("grafik_fpd_bpd_ghosh", height = "520px"))
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("lightbulb"), " Interpretasi Otomatis [Berdasarkan Inverse Ghosh]"),
          div(class = "card-body-inner", uiOutput("narasi_ui_ghosh"))
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("columns"), " Agregat"),
          div(
            class = "card-body-inner",
            uiOutput("kuadran_note_ui"),
            tags$p(style = "margin-top:1rem;", tags$b("Agregat per Baris (Output)")),
            DTOutput("tbl_kuadran"),
            tags$p(style = "margin-top:1rem;", tags$b("Agregat per Kolom (input)")),
            DTOutput("tbl_agregat")
          )
        ),
        div(
          class = "content-card",
          div(class = "card-title", icon("users"), " Labor Effect (Efek Tenaga Kerja)"),
          div(
            class = "card-body-inner",
            fileInput(
              "file_labor",
              "Data Pekerja per Sektor",
              accept = c(".xlsx", ".xls"),
              multiple = TRUE
            ),
            uiOutput("labor_info_ui"),
            uiOutput("labor_btn_ui"),
            DTOutput("tbl_labor")
          )
        ),
        div(
          class = "status-box status-info",
          icon("info-circle"),
          div("Unduhan hasil analisis tersedia pada menu ", tags$b("Output"), ".")
        )
      ),
      div(
        id = "panel-output",
        class = "main-panel",
        div(class = "page-header", h3("Output")),
        uiOutput("output_download_info"),
        div(
          style = "display:grid; grid-template-columns: 1fr 1fr; gap:1.5rem;",
          div(
            class = "content-card",
            div(class = "card-title", icon("file-excel"), " Hasil Updating I-O"),
            div(
              class = "card-body-inner",
              downloadButton(
                "btn_download",
                "Unduh Hasil Updating I-O (.xlsx)",
                class = "btn-download-result",
                icon = icon("file-excel")
              )
            )
          ),
          div(
            class = "content-card",
            div(class = "card-title", icon("file-export"), " Hasil Analisis"),
            div(
              class = "card-body-inner",
              downloadButton(
                "btn_download_full",
                "Unduh Hasil Analisis (.xlsx)",
                class = "btn-download-result",
                icon = icon("file-export")
              )
            )
          )
        )
      ),
      div(
        id = "panel-template",
        class = "main-panel",
        div(class = "page-header", h3("Unduh Template")),
        div(
          class = "content-card",
          div(class = "card-title", icon("file-excel"), " Template Berkas Input"),
          div(
            class = "card-body-inner",
            div(
              style = "display:grid; grid-template-columns: 1fr 1fr; gap:1rem;",
              div(
                class = "template-card",
                div(class = "template-icon", icon("table")),
                div(
                  class = "template-info",
                  h6("Tabel I-O — 52 Sektor"),
                  p("Matriks 52x52, kode I-01 s.d. I-52.")
                ),
                downloadButton(
                  "dl_template_io",
                  "Unduh",
                  class = "btn-template",
                  icon = icon("download")
                )
              ),
              div(
                class = "template-card",
                div(class = "template-icon", icon("bullseye")),
                div(
                  class = "template-info",
                  h6("Target — 52 Sektor"),
                  p("Total Input & Output per sektor.")
                ),
                downloadButton(
                  "dl_template_target",
                  "Unduh",
                  class = "btn-template",
                  icon = icon("download")
                )
              ),
              div(
                class = "template-card",
                div(class = "template-icon", icon("table")),
                div(
                  class = "template-info",
                  h6("Tabel I-O — 17 Sektor"),
                  p("Matriks 17x17, kode A s.d. R,S,T,U.")
                ),
                downloadButton(
                  "dl_template_io_17",
                  "Unduh",
                  class = "btn-template",
                  icon = icon("download")
                )
              ),
              div(
                class = "template-card",
                div(class = "template-icon", icon("bullseye")),
                div(
                  class = "template-info",
                  h6("Target — 17 Sektor"),
                  p("Total Input & Output per sektor agregat.")
                ),
                downloadButton(
                  "dl_template_target_17",
                  "Unduh",
                  class = "btn-template",
                  icon = icon("download")
                )
              ),
              div(
                class = "template-card",
                div(class = "template-icon", icon("users")),
                div(
                  class = "template-info",
                  h6("Tenaga Kerja — 17 Sektor"),
                  p("Kolom Kode (A-Q, MN, RSTU) + kolom jumlah pekerja.")
                ),
                downloadButton(
                  "dl_template_labor",
                  "Unduh",
                  class = "btn-template",
                  icon = icon("download")
                )
              ),
              div(
                class = "template-card",
                div(class = "template-icon", icon("users")),
                div(
                  class = "template-info",
                  h6("Tenaga Kerja — 52 Sektor"),
                  p("Kolom Kode (I-01 s.d. I-52) + kolom jumlah pekerja.")
                ),
                downloadButton(
                  "dl_template_labor_52",
                  "Unduh",
                  class = "btn-template",
                  icon = icon("download")
                )
              )
            )
          )
        )
      )
    )
  )
)

### =============================================================================
### SERVER
### =============================================================================
server <- function(input, output, session) {
  hasil <- reactiveVal(NULL)
  data_langsung <- reactiveVal(NULL)
  data_ipfp <- reactiveVal(NULL)
  sumber_aktif <- reactiveVal(NULL)
  status_msg <- reactiveVal(NULL)
  labor_data <- reactiveVal(NULL)
  labor_msg <- reactiveVal(NULL)
  labor_result <- reactiveVal(NULL)
  agregasi_17_cache <- reactiveVal(NULL)
  
  set_sumber_aktif <- function(src, silent = FALSE) {
    if (identical(src, "langsung")) {
      d <- data_langsung()
      if (is.null(d)) return(FALSE)
      hasil(d)
      sumber_aktif("langsung")
      agregasi_17_cache(NULL)
      labor_result(NULL)
      if (is_base_52(d)) {
        tryCatch(
          updateRadioButtons(session, "versi_tabel", selected = "52"),
          error = function(e) NULL
        )
      }
      if (!silent) {
        status_msg(list(
          type = "info",
          text = "Sumber data analisis aktif: Data Upload Langsung."
        ))
      }
      return(TRUE)
    }
    if (identical(src, "ipfp")) {
      d <- data_ipfp()
      if (is.null(d)) return(FALSE)
      hasil(d)
      sumber_aktif("ipfp")
      agregasi_17_cache(NULL)
      labor_result(NULL)
      if (is_base_52(d)) {
        tryCatch(
          updateRadioButtons(session, "versi_tabel", selected = "52"),
          error = function(e) NULL
        )
      }
      if (!silent) {
        status_msg(list(
          type = "info",
          text = "Sumber data analisis aktif: Hasil Updating IPFP."
        ))
      }
      return(TRUE)
    }
    FALSE
  }
  
  observeEvent(input$file_analysis, {
    req(input$file_analysis)
    session$sendCustomMessage("showOverlay", TRUE)
    session$sendCustomMessage(
      "updateProgress",
      list(percent = 5, message = "Memeriksa Data untuk Analisis...")
    )
    res <- tryCatch({
      session$sendCustomMessage(
        "updateProgress",
        list(percent = 15, message = "Mencari sheet Tabel I-O...")
      )
      loc <- find_base_sheet_in_file(input$file_analysis$datapath)
      session$sendCustomMessage(
        "updateProgress",
        list(percent = 35, message = "Membaca tabel I-O siap analisis...")
      )
      base <- read_base_io(loc$path, sheet = loc$sheet)
      session$sendCustomMessage(
        "updateProgress",
        list(percent = 75, message = "Menyiapkan objek analisis...")
      )
      make_analysis_object_from_base(base, "langsung")
    }, error = function(e) e)
    session$sendCustomMessage(
      "updateProgress",
      list(percent = 100, message = "Selesai.")
    )
    Sys.sleep(0.25)
    session$sendCustomMessage("hideOverlay", TRUE)
    if (inherits(res, "error")) {
      status_msg(list(
        type = "error",
        text = paste0("Data untuk Analisis tidak valid: ", conditionMessage(res))
      ))
      showNotification(conditionMessage(res), type = "error", duration = 10)
      return()
    }
    data_langsung(res)
    set_sumber_aktif("langsung", silent = TRUE)
    status_msg(list(
      type = "info",
      text = sprintf(
        "Data untuk Analisis berhasil dibaca (%d sektor). Silakan periksa Sumber Data Analisis dan Pilih Level Sektor pada kartu di bawah Jalur 2, kemudian buka menu Analisis.",
        length(res$codes)
      )
    ))
    showNotification("Data analisis langsung siap digunakan.", type = "message", duration = 5)
  })
  
  output$sumber_data_ui <- renderUI({
    ada_langsung <- !is.null(data_langsung())
    ada_ipfp <- !is.null(data_ipfp())
    if (!ada_langsung && !ada_ipfp) {
      return(
        div(
          class = "status-box status-info",
          icon("info-circle"),
          div(
            "Belum ada data untuk dianalisis. ",
            "Silakan unggah Data untuk Analisis atau lakukan Updating IO terlebih dahulu."
          )
        )
      )
    }
    choices <- character(0)
    if (ada_langsung) {
      choices <- c(choices, "Data Upload Langsung" = "langsung")
    }
    if (ada_ipfp) {
      choices <- c(choices, "Hasil Updating IPFP" = "ipfp")
    }
    active <- sumber_aktif()
    if (is.null(active) || !(active %in% choices)) {
      active <- unname(choices[1])
    }
    status_txt <- if (identical(active, "ipfp")) {
      "Sumber Data Analisis: Hasil Updating IPFP"
    } else {
      "Sumber Data Analisis: Data Upload Langsung"
    }
    div(
      class = "content-card",
      div(class = "card-title", icon("database"), " Sumber Data Analisis"),
      div(
        class = "card-body-inner",
        div(
          class = "status-box status-info",
          icon("check-circle"),
          div(status_txt)
        ),
        radioButtons(
          "pilih_sumber",
          "Pilih sumber data aktif:",
          choices = choices,
          selected = active,
          inline = TRUE
        )
      )
    )
  })
  
  observeEvent(input$pilih_sumber, {
    req(input$pilih_sumber)
    if (identical(input$pilih_sumber, sumber_aktif())) return()
    set_sumber_aktif(input$pilih_sumber)
  })
  
  observeEvent(sumber_aktif(), {
    session$sendCustomMessage(
      "setHasilAccess",
      !identical(sumber_aktif(), "langsung")
    )
  }, ignoreInit = FALSE)
  
  output$update_validation_ui <- renderUI({
    ready <- !is.null(input$file_base) && !is.null(input$file_target)
    if (!ready) {
      return(
        div(
          class = "status-box status-info",
          style = "margin-top:0; margin-bottom:1.25rem;",
          icon("info-circle"),
          div(
            "Untuk melakukan Updating IO, silakan upload Matriks IO Tahun Dasar dan Data Target terlebih dahulu."
          )
        )
      )
    }
    NULL
  })
  
  output$btn_update_ui <- renderUI({
    ready <- !is.null(input$file_base) && !is.null(input$file_target)
    if (ready) {
      actionButton(
        "btn_process",
        "Proses Updating IO",
        class = "btn-process",
        icon = icon("play")
      )
    } else {
      tags$button(
        class = "btn-process",
        disabled = "disabled",
        icon("play"),
        "Proses Updating IO"
      )
    }
  })
  
  versi_aktif <- reactive({
    h <- hasil()
    if (is.null(h)) return(NULL)
    if (!is_base_52(h)) return("17")
    v <- input$versi_tabel
    if (is.null(v) || !v %in% c("52", "17")) v <- "52"
    v
  })
  
  output$status_ui <- renderUI({
    s <- status_msg()
    if (is.null(s)) return(NULL)
    icn <- switch(
      s$type,
      "info" = icon("info-circle"),
      "warning" = icon("exclamation-triangle"),
      "error" = icon("times-circle"),
      icon("info-circle")
    )
    div(class = paste0("status-box status-", s$type), icn, div(s$text))
  })
  
  output$hasil_info_ui <- renderUI({
    h <- hasil()
    if (is.null(h)) {
      return(
        div(
          class = "status-box status-info",
          icon("info-circle"),
          div(
            "Belum ada data untuk dianalisis. ",
            "Silakan unggah Data untuk Analisis atau lakukan Updating IO terlebih dahulu."
          )
        )
      )
    }
    src <- sumber_aktif()
    if (is.null(src)) src <- h$source
    if (is.null(src)) src <- "langsung"
    if (identical(src, "ipfp")) {
      res <- h$res
      div(
        class = "info-grid",
        style = "margin-bottom:1.5rem;",
        div(
          class = "info-item",
          div(class = "label", "Sumber Data"),
          div(class = "value", style = "font-size:0.9rem;", "Hasil Updating IPFP")
        ),
        div(
          class = "info-item",
          div(class = "label", "Ukuran IPFP"),
          div(class = "value", paste0(length(h$codes), " x ", length(h$codes)))
        ),
        div(
          class = "info-item",
          div(class = "label", "Status"),
          div(
            class = "value",
            style = "color:#10B981;",
            ifelse(res$converged, "Konvergen", "Belum Konvergen")
          )
        ),
        div(
          class = "info-item",
          div(class = "label", "Iterasi"),
          div(class = "value", res$iterations)
        ),
        div(
          class = "info-item",
          div(class = "label", "Selisih Maks"),
          div(class = "value", format(res$max_diff, scientific = TRUE, digits = 3))
        ),
        div(
          class = "info-item",
          div(class = "label", "Faktor Skala v*"),
          div(class = "value", format(res$scale_factor, digits = 5))
        )
      )
    } else {
      div(
        class = "info-grid",
        style = "margin-bottom:1.5rem;",
        div(
          class = "info-item",
          div(class = "label", "Sumber Data"),
          div(class = "value", style = "font-size:0.9rem;", "Data Upload Langsung")
        ),
        div(
          class = "info-item",
          div(class = "label", "Ukuran Tabel"),
          div(class = "value", paste0(length(h$codes), " x ", length(h$codes)))
        ),
        div(
          class = "info-item",
          div(class = "label", "Mode"),
          div(class = "value", "Analisis Langsung")
        ),
        div(
          class = "info-item",
          div(class = "label", "Status"),
          div(class = "value", style = "color:#10B981;", "Siap Dianalisis")
        )
      )
    }
  })
  
  output$versi_ui <- renderUI({
    h <- hasil()
    if (is.null(h)) return(NULL)
    if (!is_base_52(h)) {
      return(
        div(
          class = "status-box status-info",
          icon("info-circle"),
          div("Data terdeteksi berukuran 17 sektor — unduhan dan analisis menggunakan versi 17 sektor.")
        )
      )
    }
    div(
      class = "content-card",
      div(class = "card-title", icon("sitemap"), " Pilih Level Sektor untuk Analisis"), # IKON-DIPERBARUI
      div(
        class = "card-body-inner",
        radioButtons(
          "versi_tabel",
          NULL,
          choices = list(
            "52 sektor" = "52",
            "17 sektor" = "17"
          ),
          selected = "52",
          inline = TRUE
        )
      )
    )
  })
  
  tabel_analisis <- reactive({
    req(hasil())
    v <- versi_aktif()
    h <- hasil()
    if (is_base_52(h) && identical(v, "17")) {
      cached <- agregasi_17_cache()
      if (!is.null(cached)) return(cached)
      agg_tbl <- build_analysis_table(h, "17")
      agregasi_17_cache(agg_tbl)
      return(agg_tbl)
    }
    build_analysis_table(h, v)
  })
  
  analisis <- reactive({
    req(tabel_analisis())
    compute_analysis(tabel_analisis())
  })
  
  kuadran <- reactive({
    req(tabel_analisis())
    compute_quadrant2(tabel_analisis())
  })
  
  observeEvent(list(input$versi_tabel, sumber_aktif()), {
    labor_result(NULL)
  }, ignoreInit = TRUE)
  
  observeEvent(input$btn_process, {
    if (is.null(input$file_base) || is.null(input$file_target)) {
      status_msg(list(
        type = "warning",
        text = "Untuk melakukan Updating IO, silakan upload Matriks IO Tahun Dasar dan Data Target terlebih dahulu."
      ))
      return()
    }
    session$sendCustomMessage("showOverlay", TRUE)
    session$sendCustomMessage(
      "updateProgress",
      list(percent = 0, message = "Memulai proses...")
    )
    Sys.sleep(0.3)
    result <- tryCatch({
      session$sendCustomMessage(
        "updateProgress",
        list(percent = 5, message = "Memeriksa berkas yang diunggah...")
      )
      Sys.sleep(0.15)
      inp <- resolve_inputs(list(input$file_base, input$file_target))
      session$sendCustomMessage(
        "updateProgress",
        list(percent = 12, message = "Membaca Tabel I-O tahun dasar...")
      )
      Sys.sleep(0.15)
      base <- read_base_io(inp$base_path, sheet = inp$base_sheet)
      session$sendCustomMessage(
        "updateProgress",
        list(percent = 22, message = "Membaca Data Tahun Target periode baru...")
      )
      Sys.sleep(0.15)
      target <- read_target_io(inp$target_path, sheet = inp$target_sheet)
      session$sendCustomMessage(
        "updateProgress",
        list(
          percent = 26,
          message = "Memeriksa kesesuaian ukuran Data Dasar dan Data Tahun Target..."
        )
      )
      Sys.sleep(0.12)
      validate_target_dimension(base, target)
      session$sendCustomMessage(
        "updateProgress",
        list(
          percent = 30,
          message = sprintf("Menyelaraskan %d kode sektor antar berkas...", base$n)
        )
      )
      Sys.sleep(0.12)
      target_aligned <- align_target_to_base(base$codes, target)
      max_iter <- input$max_iter
      tol_value <- input$tol_value
      if (is.null(max_iter) || is.na(max_iter) || max_iter < 1) max_iter <- 1000
      if (is.null(tol_value) || is.na(tol_value) || tol_value <= 0) tol_value <- 1e-6
      res <- compute_ipfp_update(base, target_aligned, max_iter, tol_value, session)
      session$sendCustomMessage(
        "updateProgress",
        list(percent = 98, message = "Menyimpan hasil...")
      )
      Sys.sleep(0.15)
      list(
        codes = base$codes,
        names = base$names,
        res = res,
        fd = base$fd,
        base_output = base$total_output,
        error = NULL,
        source = "ipfp"
      )
    }, error = function(e) list(error = conditionMessage(e)))
    session$sendCustomMessage(
      "updateProgress",
      list(percent = 100, message = "Selesai!")
    )
    Sys.sleep(0.4)
    session$sendCustomMessage("hideOverlay", TRUE)
    if (!is.null(result$error)) {
      status_msg(list(
        type = "error",
        text = paste0("Terjadi kesalahan: ", result$error)
      ))
      showNotification(result$error, type = "error", duration = 10)
      return(invisible(NULL))
    }
    data_ipfp(result)
    set_sumber_aktif("ipfp", silent = TRUE)
    res <- result$res
    if (res$converged) {
      status_msg(list(
        type = "info",
        text = sprintf(
          "IPFP berhasil pada %d sektor, konvergen dalam %d iterasi (selisih maks %.2e). Silakan tentukan sumber data dan level sektor pada kartu di bawah Jalur 2 sebelum membuka menu Analisis.",
          length(result$codes),
          res$iterations,
          res$max_diff
        )
      ))
    } else {
      status_msg(list(
        type = "warning",
        text = sprintf(
          "Proses selesai tetapi belum konvergen penuh (%d iterasi, selisih maks %.2e). Coba naikkan maksimum iterasi.",
          res$iterations,
          res$max_diff
        )
      ))
    }
    showNotification("IPFP selesai.", type = "message", duration = 5)
  })
  
  output$tbl_ringkasan <- renderDT({
    req(hasil(), tabel_analisis())
    tbl <- tabel_analisis()
    n_show <- suppressWarnings(as.integer(input$rows_ringkasan))
    if (is.na(n_show)) n_show <- 25
    df <- data.frame(
      Kode = tbl$codes,
      `Nama Sektor` = tbl$names,
      `Total Output` = round(tbl$output, 2),
      `Total Input` = round(tbl$input, 2),
      `Konsumsi Akhir` = round(tbl$konsumsi, 2),
      `NTB` = round(tbl$ntb, 2),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    datatable(
      df,
      rownames = FALSE,
      options = list(
        pageLength = n_show,
        scrollX = TRUE,
        autoWidth = FALSE,
        dom = "ftip"
      ),
      class = "stripe hover"
    ) %>%
      formatCurrency(
        columns = c("Total Output", "Total Input", "Konsumsi Akhir", "NTB"),
        currency = "",
        interval = 3,
        mark = ",",
        digits = 2
      )
  })
  
  output$tbl_matriks <- renderDT({
    req(hasil(), tabel_analisis())
    tbl <- tabel_analisis()
    codes <- tbl$codes
    n_show <- suppressWarnings(as.integer(input$rows_matriks))
    if (is.na(n_show)) n_show <- 10
    Z_disp <- round(tbl$Z, 2)
    rownames(Z_disp) <- codes
    colnames(Z_disp) <- codes
    datatable(
      Z_disp,
      options = list(
        scrollX = TRUE,
        autoWidth = FALSE,
        pageLength = n_show,
        dom = "ftip"
      ),
      class = "stripe hover"
    ) %>%
      formatCurrency(
        columns = codes,
        currency = "",
        interval = 3,
        mark = ",",
        digits = 2
      )
  })
  
  output$output_download_info <- renderUI({
    h <- hasil()
    if (is.null(h)) {
      return(
        div(
          class = "status-box status-warning",
          icon("exclamation-triangle"),
          div(
            "Belum ada data untuk dianalisis. ",
            "Unggah Data untuk Analisis atau jalankan Updating IO terlebih dahulu."
          )
        )
      )
    }
    tbl <- tabel_analisis()
    v <- versi_aktif()
    sumber <- if (!is.null(tbl)) tbl$sumber else "Belum tersedia"
    src_label <- if (identical(sumber_aktif(), "ipfp")) {
      "Hasil Updating IPFP"
    } else {
      "Data Upload Langsung"
    }
    div(
      class = "info-grid",
      style = "margin-bottom:1.5rem;",
      div(
        class = "info-item",
        div(class = "label", "Sumber Data Aktif"),
        div(class = "value", style = "font-size:0.9rem;", src_label)
      ),
      div(
        class = "info-item",
        div(class = "label", "Ukuran Tabel Asli"),
        div(class = "value", paste0(length(h$codes), " x ", length(h$codes)))
      ),
      div(
        class = "info-item",
        div(class = "label", "Level Output Aktif"),
        div(class = "value", paste0(v, " sektor"))
      ),
      div(
        class = "info-item",
        div(class = "label", "Sumber Analisis"),
        div(class = "value", style = "font-size:.9rem;", sumber)
      )
    )
  })
  
  output$btn_download <- downloadHandler(
    filename = function() {
      v <- versi_aktif()
      paste0(
        "Tabel_IO_Aktif_",
        ifelse(is.null(v), "NA", v),
        "_",
        format(Sys.Date(), "%Y%m%d"),
        ".xlsx"
      )
    },
    content = function(file) {
      req(hasil())
      tbl_io <- tabel_analisis()
      if (is.null(tbl_io)) {
        h <- hasil()
        res <- h$res
        tbl_io <- list(
          codes = h$codes,
          names = h$names,
          Z = res$Z_new,
          konsumsi = res$konsumsi_akhir_new,
          output = res$total_output_new,
          input = res$total_input_new,
          ntb = res$ntb_new,
          impor_ln = res$impor_ln_new,
          impor_ap = res$impor_ap_new,
          fd_new = list()
        )
        q2_io <- compute_quadrant2(tbl_io)
      } else {
        q2_io <- kuadran()
      }
      wb <- build_update_io_only_workbook(hasil(), tbl_io, q2_io)
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  output$analisis_status_ui <- renderUI({
    h <- hasil()
    tbl <- tabel_analisis()
    if (is.null(h)) {
      return(
        div(
          class = "status-box status-info",
          icon("info-circle"),
          div(
            "Belum ada data untuk dianalisis. ",
            "Silakan unggah ",
            tags$b("Data untuk Analisis"),
            " pada menu Input Data, atau lakukan ",
            tags$b("Updating IO"),
            " terlebih dahulu."
          )
        )
      )
    }
    if (is.null(tbl)) {
      return(
        div(
          class = "status-box status-warning",
          icon("exclamation-triangle"),
          div("Analisis belum tersedia. Pastikan data valid dan level sektor telah ditentukan.")
        )
      )
    }
    an <- analisis()
    src_label <- if (identical(sumber_aktif(), "ipfp")) {
      "Hasil Updating IPFP"
    } else {
      "Data Upload Langsung"
    }
    div(
      class = "info-grid",
      style = "margin-bottom:1.5rem;",
      div(
        class = "info-item",
        div(class = "label", "Sumber Data"),
        div(class = "value", style = "font-size:0.9rem;", src_label)
      ),
      div(
        class = "info-item",
        div(class = "label", "Sektor Analisis"),
        div(class = "value", length(tbl$codes))
      ),
      div(
        class = "info-item",
        div(class = "label", "Sumber Tabel"),
        div(class = "value", style = "font-size:0.9rem;", tbl$sumber)
      ),
      div(
        class = "info-item",
        div(class = "label", "Inverse Leontief (I-A)"),
        div(
          class = "value",
          style = ifelse(an$okA, "color:#10B981;", "color:#DC2626;"),
          ifelse(an$okA, "Valid", "Singular")
        )
      ),
      div(
        class = "info-item",
        div(class = "label", "Inverse Ghosh (I-R)"),
        div(
          class = "value",
          style = ifelse(an$okB, "color:#10B981;", "color:#DC2626;"),
          ifelse(an$okB, "Valid", "Singular")
        )
      )
    )
  })
  
  dt_matrix <- function(M, codes, digits = 6) {
    Md <- round(as.matrix(M), digits)
    rownames(Md) <- codes
    colnames(Md) <- codes
    datatable(
      Md,
      options = list(
        scrollX = TRUE,
        autoWidth = FALSE,
        pageLength = 10,
        dom = "ftip"
      ),
      class = "stripe hover"
    )
  }
  
  output$tbl_A <- renderDT({
    req(analisis(), tabel_analisis())
    dt_matrix(analisis()$A, tabel_analisis()$codes)
  })
  output$tbl_invA <- renderDT({
    req(analisis(), tabel_analisis())
    dt_matrix(analisis()$invA, tabel_analisis()$codes)
  })
  output$tbl_B <- renderDT({
    req(analisis(), tabel_analisis())
    dt_matrix(analisis()$B, tabel_analisis()$codes)
  })
  output$tbl_invB <- renderDT({
    req(analisis(), tabel_analisis())
    dt_matrix(analisis()$invB, tabel_analisis()$codes)
  })
  
  output$tbl_multiplier_leontief <- renderDT({
    req(analisis(), tabel_analisis())
    an <- analisis()
    tbl <- tabel_analisis()
    df <- data.frame(
      Kode = tbl$codes,
      `Nama Sektor` = tbl$names,
      `Koefisien Nilai Tambah (%)` = round(an$phi * 100, 2),
      `Output Multiplier Leontief` = round(an$OM_L, 4),
      `Nilai Tambah Multiplier Leontief` = round(an$VAM_L, 4),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    datatable(
      df,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = FALSE,
        pageLength = 17,
        dom = "ftip"
      ),
      class = "stripe hover"
    )
  })
  
  output$tbl_linkage_leontief <- renderDT({
    req(analisis(), tabel_analisis())
    an <- analisis()
    tbl <- tabel_analisis()
    df <- data.frame(
      Kode = tbl$codes,
      `Nama Sektor` = tbl$names,
      `Forward Linkage` = round(an$FL_L, 4),
      `Backward Linkage` = round(an$BL_L, 4),
      `FPD` = round(an$FPD_L, 4),
      `BPD` = round(an$BPD_L, 4),
      `FPD ket` = ifelse(an$FPD_L > 1, ">1", "<1"),
      `BPD ket` = ifelse(an$BPD_L > 1, ">1", "<1"),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    datatable(
      df,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = FALSE,
        pageLength = 17,
        dom = "ftip"
      ),
      class = "stripe hover"
    )
  })
  
  output$narasi_ui_leontief <- renderUI({
    req(analisis(), tabel_analisis())
    tbl <- tabel_analisis()
    tags$ul(
      class = "narasi-list",
      lapply(
        build_narratives_model(
          analisis(),
          tbl$codes,
          tbl$names,
          model = "leontief"
        ),
        tags$li
      )
    )
  })
  
  output$grafik_fpd_bpd_leontief <- renderPlotly({
    req(analisis(), tabel_analisis())
    an  <- analisis()
    tbl <- tabel_analisis()
    df  <- make_fpd_bpd_df(an$FPD_L, an$BPD_L, tbl)
    validate(
      need(
        nrow(df) > 0,
        "Grafik tidak dapat ditampilkan: nilai FPD/BPD Leontief tidak tersedia (matriks singular)."
      )
    )
    build_fpd_bpd_plot(df, "Inverse Leontief")
  })
  
  output$grafik_fpd_bpd_ghosh <- renderPlotly({
    req(analisis(), tabel_analisis())
    an  <- analisis()
    tbl <- tabel_analisis()
    df  <- make_fpd_bpd_df(an$FPD_G, an$BPD_G, tbl)
    validate(
      need(
        nrow(df) > 0,
        "Grafik tidak dapat ditampilkan: nilai FPD/BPD Ghosh tidak tersedia (matriks singular)."
      )
    )
    build_fpd_bpd_plot(df, "Inverse Ghosh")
  })
  
  output$tbl_multiplier_ghosh <- renderDT({
    req(analisis(), tabel_analisis())
    an <- analisis()
    tbl <- tabel_analisis()
    df <- data.frame(
      Kode = tbl$codes,
      `Nama Sektor` = tbl$names,
      `Koefisien Nilai Tambah (%)` = round(an$phi * 100, 2),
      `Output Multiplier Ghosh` = round(an$OM_G, 4),
      `Nilai Tambah Multiplier Ghosh` = round(an$VAM_G, 4),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    datatable(
      df,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = FALSE,
        pageLength = 17,
        dom = "ftip"
      ),
      class = "stripe hover"
    )
  })
  
  output$tbl_linkage_ghosh <- renderDT({
    req(analisis(), tabel_analisis())
    an <- analisis()
    tbl <- tabel_analisis()
    df <- data.frame(
      Kode = tbl$codes,
      `Nama Sektor` = tbl$names,
      `Forward Linkage` = round(an$FL_G, 4),
      `Backward Linkage` = round(an$BL_G, 4),
      `FPD` = round(an$FPD_G, 4),
      `BPD` = round(an$BPD_G, 4),
      `FPD ket` = ifelse(an$FPD_G > 1, ">1", "<1"),
      `BPD ket` = ifelse(an$BPD_G > 1, ">1", "<1"),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    datatable(
      df,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = FALSE,
        pageLength = 17,
        dom = "ftip"
      ),
      class = "stripe hover"
    )
  })
  
  output$narasi_ui_ghosh <- renderUI({
    req(analisis(), tabel_analisis())
    tbl <- tabel_analisis()
    tags$ul(
      class = "narasi-list",
      lapply(
        build_narratives_model(
          analisis(),
          tbl$codes,
          tbl$names,
          model = "ghosh"
        ),
        tags$li
      )
    )
  })
  
  output$tbl_kuadran <- renderDT({
    req(kuadran(), tabel_analisis())
    q2 <- kuadran()
    tbl <- tabel_analisis()
    rnd <- function(x) {
      round(as.numeric(if (is.null(x)) rep(NA_real_, length(tbl$codes)) else x), 2)
    }
    df <- data.frame(
      Kode = tbl$codes,
      `1800` = rnd(q2$k1800),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    if (isTRUE(q2$has_fd)) {
      df_fd <- data.frame(
        `3011` = rnd(q2$k3011),
        `3012` = rnd(q2$k3012),
        `3020` = rnd(q2$k3020),
        `3030` = rnd(q2$k3030),
        `3041` = rnd(q2$k3041),
        `3071` = rnd(q2$k3071),
        `3072` = rnd(q2$k3072),
        `3080` = rnd(q2$k3080),
        check.names = FALSE,
        stringsAsFactors = FALSE
      )
      df <- cbind(df, df_fd)
    }
    df_tail <- data.frame(
      `3090` = rnd(q2$k3090),
      `3100` = rnd(q2$k3100),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    df <- cbind(df, df_tail)
    datatable(
      df,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = FALSE,
        pageLength = 17,
        dom = "ftip"
      ),
      class = "stripe hover"
    ) %>%
      formatCurrency(
        columns = setdiff(names(df), "Kode"),
        currency = "",
        interval = 3,
        mark = ",",
        digits = 2
      )
  })
  
  output$tbl_agregat <- renderDT({
    req(kuadran(), tabel_analisis())
    q2 <- kuadran()
    tbl <- tabel_analisis()
    M <- rbind(
      q2$k190d,
      tbl$impor_ln,
      tbl$impor_ap,
      q2$k1900,
      tbl$ntb,
      tbl$input
    )
    df <- as.data.frame(round(M, 2))
    colnames(df) <- tbl$codes
    df <- cbind(
      Kode = c("190d", "2000", "2001", "1900", "2090", "2100"),
      Keterangan = c(
        "Input Antara Domestik",
        "Impor Luar Negeri",
        "Impor Antar Provinsi",
        "Total Input Antara",
        "Nilai Tambah Bruto",
        "Total Input"
      ),
      df
    )
    datatable(
      df,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = FALSE,
        dom = "ftip"
      ),
      class = "stripe hover"
    )
  })
  
  observeEvent(input$file_labor, {
    labor_result(NULL)
    if (is.null(hasil()) || is.null(tabel_analisis())) {
      labor_data(NULL)
      labor_msg(list(
        type = "warning",
        text = "Belum ada data analisis. Unggah Data untuk Analisis atau jalankan Updating IO terlebih dahulu sebelum mengunggah data tenaga kerja."
      ))
      return()
    }
    res <- tryCatch(
      read_labor_data_multi(input$file_labor$datapath, tabel_analisis()$codes),
      error = function(e) e
    )
    if (inherits(res, "error")) {
      labor_data(NULL)
      labor_msg(list(
        type = "error",
        text = paste0("Berkas tenaga kerja: ", conditionMessage(res))
      ))
    } else {
      labor_data(res$values)
      col_desc <- if (isTRUE(nchar(res$info$column) > 0)) {
        res$info$column
      } else {
        sprintf("kolom #%d", res$info$col_index)
      }
    }
  })
  
  output$labor_info_ui <- renderUI({
    m <- labor_msg()
    if (is.null(m)) return(NULL)
    icn <- switch(
      m$type,
      "info" = icon("info-circle"),
      "warning" = icon("exclamation-triangle"),
      "error" = icon("times-circle"),
      icon("info-circle")
    )
    div(class = paste0("status-box status-", m$type), icn, div(m$text))
  })
  
  output$labor_btn_ui <- renderUI({
    if (is.null(hasil()) || is.null(tabel_analisis())) return(NULL)
    if (is.null(labor_data())) {
      return(
        div(
          class = "status-box status-info",
          icon("info-circle"),
          div("Belum ada data tenaga kerja yang diunggah.")
        )
      )
    }
    actionButton(
      "btn_labor",
      "Jalankan Labor Effect",
      class = "btn-process",
      icon = icon("users")
    )
  })
  
  observeEvent(input$btn_labor, {
    req(hasil(), tabel_analisis(), labor_data())
    tbl <- tabel_analisis()
    if (length(labor_data()) != length(tbl$codes)) {
      labor_result(NULL)
      showNotification(
        sprintf(
          "Data tenaga kerja harus memiliki %d sektor sesuai level analisis aktif. Silakan unggah ulang data tenaga kerja yang sesuai.",
          length(tbl$codes)
        ),
        type = "error",
        duration = 10
      )
      return()
    }
    an <- analisis()
    if (!an$okA) {
      showNotification("Invers (I-A) singular — Labor Effect tidak dapat dihitung.", type = "error")
      return()
    }
    labor_result(compute_labor(labor_data(), tbl, an$invA))
    showNotification("Labor Effect berhasil dihitung.", type = "message", duration = 5)
  })
  
  output$tbl_labor <- renderDT({
    req(labor_result(), tabel_analisis())
    lab <- labor_result()
    tbl <- tabel_analisis()
    df <- data.frame(
      Kode = tbl$codes,
      `Nama Sektor` = tbl$names,
      Pekerja = lab$L,
      Produktivitas = round(lab$prod, 4),
      `Koefisien Tenaga Kerja` = round(lab$l, 6),
      `Labor Multiplier` = round(lab$LM, 4),
      `Employment Multiplier` = round(lab$EM, 6),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    datatable(
      df,
      rownames = FALSE,
      options = list(
        scrollX = TRUE,
        autoWidth = FALSE,
        pageLength = 17,
        dom = "ftip"
      ),
      class = "stripe hover"
    )
  })
  
  output$btn_download_full <- downloadHandler(
    filename = function() {
      v <- versi_aktif()
      paste0(
        "Hasil_Analisis_IO_",
        ifelse(is.null(v), "NA", v),
        "_",
        format(Sys.Date(), "%Y%m%d"),
        ".xlsx"
      )
    },
    content = function(file) {
      req(hasil(), tabel_analisis(), analisis(), kuadran())
      df_fpd_leontief <- make_fpd_bpd_df(
        analisis()$FPD_L,
        analisis()$BPD_L,
        tabel_analisis()
      )
      df_fpd_ghosh <- make_fpd_bpd_df(
        analisis()$FPD_G,
        analisis()$BPD_G,
        tabel_analisis()
      )
      images <- list(
        leontief = save_fpd_bpd_chart(
          df_fpd_leontief,
          "Inverse Leontief",
          "fpd_bpd_leontief",
          width_cm = 18,
          height_cm = 12
        ),
        ghosh = save_fpd_bpd_chart(
          df_fpd_ghosh,
          "Inverse Ghosh",
          "fpd_bpd_ghosh",
          width_cm = 18,
          height_cm = 12
        )
      )
      wb <- build_template_output_workbook(
        hasil(),
        tabel_analisis(),
        analisis(),
        kuadran(),
        labor_result(),
        images = images
      )
      saveWorkbook(wb, file, overwrite = TRUE)
      for (f in unlist(images)) {
        if (is.character(f) && file.exists(f)) {
          unlink(f)
        }
      }
    }
  )
  
  output$dl_template_io <- downloadHandler(
    filename = function() "Template_Input_Matriks_52x52.xlsx",
    content = function(file) {
      src <- file.path(getwd(), "Template Input Matriks 52x52.xlsx")
      if (file.exists(src)) {
        file.copy(src, file, overwrite = TRUE)
      } else {
        wb <- generate_template_io(CODES_52, NAMES_52)
        saveWorkbook(wb, file, overwrite = TRUE)
      }
    }
  )
  output$dl_template_target <- downloadHandler(
    filename = function() "Template_Tahun_Target_52x52.xlsx",
    content = function(file) {
      src <- file.path(getwd(), "Template Tahun Target - 52x52.xlsx")
      if (file.exists(src)) {
        file.copy(src, file, overwrite = TRUE)
      } else {
        wb <- generate_template_target(CODES_52, NAMES_52)
        saveWorkbook(wb, file, overwrite = TRUE)
      }
    }
  )
  output$dl_template_io_17 <- downloadHandler(
    filename = function() "Template_Input_Matriks_17x17.xlsx",
    content = function(file) {
      src <- file.path(getwd(), "Template Input Matriks 17x17.xlsx")
      if (file.exists(src)) {
        file.copy(src, file, overwrite = TRUE)
      } else {
        wb <- generate_template_io(CODES_17, NAMES_17)
        saveWorkbook(wb, file, overwrite = TRUE)
      }
    }
  )
  output$dl_template_target_17 <- downloadHandler(
    filename = function() "Template_Tahun_Target_17x17.xlsx",
    content = function(file) {
      src <- file.path(getwd(), "Template Tahun Target - 17x17.xlsx")
      if (file.exists(src)) {
        file.copy(src, file, overwrite = TRUE)
      } else {
        wb <- generate_template_target(CODES_17, NAMES_17)
        saveWorkbook(wb, file, overwrite = TRUE)
      }
    }
  )
  output$dl_template_labor <- downloadHandler(
    filename = function() "Template_Data_Pekerja_17x17.xlsx",
    content = function(file) {
      src <- file.path(getwd(), "Template Data Pekerja 17x17.xlsx")
      if (file.exists(src)) {
        file.copy(src, file, overwrite = TRUE)
      } else {
        wb <- generate_template_labor(CODES_17)
        saveWorkbook(wb, file, overwrite = TRUE)
      }
    }
  )
  output$dl_template_labor_52 <- downloadHandler(
    filename = function() "Template_Data_Pekerja_52x52.xlsx",
    content = function(file) {
      src <- file.path(getwd(), "Template Data Pekerja 52x52.xlsx")
      if (file.exists(src)) {
        file.copy(src, file, overwrite = TRUE)
      } else {
        wb <- generate_template_labor(CODES_52)
        saveWorkbook(wb, file, overwrite = TRUE)
      }
    }
  )
}

### =============================================================================
### JALANKAN APLIKASI
### =============================================================================
shinyApp(ui = ui, server = server)