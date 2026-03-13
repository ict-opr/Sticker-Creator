import re
import os
import sys
import csv
import unicodedata
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# -----------------------------
# Helper: base dir (supports PyInstaller)
# -----------------------------
def get_base_dir():
	if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
		return sys._MEIPASS
	return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()

# -----------------------------
# Behavior toggles
# -----------------------------
# Controls what value is shown for Box No.
# - "per_article": restart at 1 for each CSV row/article (current/original behavior)
# - "per_order": continue counting across all boxes that belong to the same Order No.
#                If Order No. is empty for a row, it falls back to per_article for that row.
CARTON_NUMBERING_MODE = "per_order"

# Controls how PDFs are written.
# - "multipage": write all labels into a single multi-page PDF
# - "separate": write one PDF file per label (old behavior)
PDF_OUTPUT_MODE = "multipage"

# -----------------------------
# Font setup (Fonts/ directory)
# -----------------------------
fonts_dir = os.path.join(BASE_DIR, "Fonts")
font_regular_path = os.path.join(fonts_dir, "Helvetica.ttf")
font_bold_path = os.path.join(fonts_dir, "Helvetica-Bold.ttf")

FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

try:
	if os.path.exists(font_regular_path):
		pdfmetrics.registerFont(TTFont("HelveticaCustom", font_regular_path))
		FONT_REGULAR = "HelveticaCustom"
	if os.path.exists(font_bold_path):
		pdfmetrics.registerFont(TTFont("HelveticaCustom-Bold", font_bold_path))
		FONT_BOLD = "HelveticaCustom-Bold"
	else:
		# If no bold provided, use regular as bold fallback
		if FONT_REGULAR == "HelveticaCustom":
			FONT_BOLD = "HelveticaCustom"
except Exception:
	# Fall back to built-ins if registration fails
	FONT_REGULAR = "Helvetica"
	FONT_BOLD = "Helvetica-Bold"

# -----------------------------
# Sanitizers / parsing
# -----------------------------
def sanitize_numeric_keep_zeros(value, default=None):
	if value is None:
		return default
	s = re.sub(r"\D", "", str(value).strip())
	return s if s != "" else default


def sanitize_any(value, default=None):
	if value is None:
		return default
	s = str(value).strip()
	return s if s != "" else default


def strip_leading_zeros_keep_zero(s):
	s = sanitize_numeric_keep_zeros(s, default="0")
	s2 = s.lstrip("0")
	return s2 if s2 != "" else "0"


def safe_filename_component(value, default=""):
	s = sanitize_any(value, default=default) or default
	s = re.sub(r'[\\/*?:"<>|]', "_", s)
	s = re.sub(r"\s+", " ", s).strip()
	return s


def truncate_component(value, max_len):
	if value is None:
		return ""
	s = str(value)
	return s if len(s) <= max_len else s[:max_len]


def normalize_for_match(s):
	# Lower + strip accents for robust header matching
	if s is None:
		return ""
	s = str(s).strip().lower()
	s = unicodedata.normalize("NFKD", s)
	s = "".join(ch for ch in s if not unicodedata.combining(ch))
	return s


def build_column_map(header_row):
	"""Build a best-effort column index map from the CSV header row.

	Supports both old and new header names.

	Columns:
	- COMPANY (optional; defaults to 'COMPANY_NAME' if missing)
	- ARTICLE / ARTIKEL
	- QUANTITY / STUCKZAHL / STUECKZAHL
	- ORDER INDEX / BATCH / CHARGEN
	- ORDER / BESTELL (optional)
	- CARTON / KARTON
	- WEIGHT / GEWICHT (optional)
	"""
	normalized = [normalize_for_match(h) for h in header_row]

	def find_idx(patterns, exclude_patterns=None):
		exclude_patterns = exclude_patterns or []
		for i, h in enumerate(normalized):
			if any(re.search(ex_p, h) for ex_p in exclude_patterns):
				continue
			for p in patterns:
				if re.search(p, h):
					return i
		return None

	order_index_idx = find_idx([
		r"^order\s*index$",
		r"^orderindex$",
		r"^batch$",
		r"^chargen(?:\s*-?\s*nr)?$",
		r"order\s*index",
		r"batch",
		r"chargen",
	])

	order_idx = find_idx(
		[
			r"^bestell(?:\s|-)?nr$",
			r"^bestellnummer$",
			r"^order(?:\s|-)?no$",
			r"^order(?:\s|-)?number$",
			r"^order$",
			r"bestell",
			r"order",
		],
		exclude_patterns=[r"order\s*index", r"^orderindex$", r"batch", r"chargen"],
	)

	return {
		"company": find_idx([r"\bcompany\b", r"\bfirma\b", r"\bunternehmen\b"]),
		"article": find_idx([r"artikel", r"item\s*no", r"article\s*no", r"\bsku\b", r"\bitem\b", r"\barticle\b"]),
		"quantity": find_idx([r"stuckzahl", r"stueckzahl", r"quantity", r"qty", r"menge"]),
		"order_index": order_index_idx,
		"order": order_idx,
		"carton": find_idx([r"karton", r"carton"]),
		"weight": find_idx([r"\bweight\b", r"\bgewicht\b"]),
	}

def get_cell(row, idx, default=""):
	if idx is None:
		return default
	if idx < 0 or idx >= len(row):
		return default
	return row[idx]

# -----------------------------
# Core: one PDF page drawing
# -----------------------------
def draw_label_page(c, company_raw, artikel_raw, qty_raw, order_index_raw, order_raw, carton_no, weight_raw):
	company = sanitize_any(company_raw, default="COMPANY_NAME")

	# Display values (no modification other than trimming)
	artikel_display = sanitize_any(artikel_raw, default="")
	order_display = sanitize_any(order_raw, default="")
	order_index_display = sanitize_any(order_index_raw, default="1")
	weight_display = sanitize_any(weight_raw, default="")

	# Quantity display: keep as number without padding
	qty_digits = sanitize_numeric_keep_zeros(qty_raw, default="0")
	quantity_display = strip_leading_zeros_keep_zero(qty_digits)

	# -----------------------------
	# Barcode value (Code 128)
	# Required format:
	# ARTICLE_NO|QUANTITY|ORDER_INDEX|CARTON_NO
	# -----------------------------
	#barcode_data = f"{artikel_display}|{qty_digits}|{order_index_display}|{carton_no}" # Optimized from ORION Format
	barcode_data = f"{artikel_display}" # SIMPLIFIED: SKU Only!

	# Safe components for filename
	page_w, page_h = A4

	# Row heights
	# [Header, Artikel, Quantity, Barcode, Order Index, Order, Carton, Weight]
	row_heights = [
		90,   # Header row for company
		70,   # Article No.
		70,   # Quantity
		150,  # Barcode
		70,   # Order Index
		70,   # Order No.
		70,   # Box No.
		70,   # Weight
	]

	# Margins & label frame
	# Make the outer border match the actual content height exactly,
	# so layout changes do not leave an empty block at the bottom.
	margin_x = 36  # ~0.5 inch
	label_w = page_w - 2 * margin_x
	label_h = sum(row_heights)
	label_x = margin_x
	label_y = (page_h - label_h) / 2

	# Borders
	c.setLineWidth(3)  # outer border
	c.rect(label_x, label_y, label_w, label_h)
	c.setLineWidth(1.5)  # inner lines

	row_top = label_y + label_h
	col1_w = label_w * 0.45
	col2_w = label_w - col1_w

	# -----------------------------
	# Header row: COMPANY (shrink-to-fit)
	# -----------------------------
	header_h = row_heights[0]
	c.rect(label_x, row_top - header_h, label_w, header_h)

	max_font = 42
	min_font = 12
	font_size = max_font
	available_w = label_w - 24
	while font_size > min_font and pdfmetrics.stringWidth(company, FONT_BOLD, font_size) > available_w:
		font_size -= 1

	c.setFont(FONT_BOLD, font_size)
	c.drawCentredString(label_x + label_w / 2, row_top - header_h / 2 - (font_size * 0.33), company)
	row_top -= header_h

	# -----------------------------
	# Row helper
	# - Main label: bold
	# - Sub label: regular
	# - Right-side value: regular
	# -----------------------------
	def draw_row(y_top, height, main_label, sub_label, value=None, shrinkable=False):
		# Left & right boxes
		c.rect(label_x, y_top - height, col1_w, height)
		c.rect(label_x + col1_w, y_top - height, col2_w, height)

		# Left-side labels
		main_label_font_size = 36
		c.setFont(FONT_BOLD, main_label_font_size)
		c.drawString(label_x + 12, y_top - 30, main_label)

		c.setFont(FONT_REGULAR, 18)
		c.drawString(label_x + 12, y_top - 54, sub_label)

		# Right-side value
		if value is not None:
			value = str(value)
			font_size_val = main_label_font_size
			if shrinkable:
				text_w = pdfmetrics.stringWidth(value, FONT_REGULAR, font_size_val)
				available_w_val = col2_w - 20
				while text_w > available_w_val and font_size_val > 8:
					font_size_val -= 1
					text_w = pdfmetrics.stringWidth(value, FONT_REGULAR, font_size_val)
			c.setFont(FONT_REGULAR, font_size_val)
			text_y = y_top - (height / 2) - (font_size_val * 0.35)
			c.drawString(label_x + col1_w + 10, text_y, value)

	# -----------------------------
	# Article No. (full, unmodified; shrink-to-fit)
	# -----------------------------
	rh = row_heights[1]
	draw_row(row_top, rh, "Article No.", "(SKU)", artikel_display, shrinkable=True)
	row_top -= rh

	# -----------------------------
	# Quantity (display WITHOUT padding)
	# -----------------------------
	rh = row_heights[2]
	draw_row(row_top, rh, "Quantity", "(Per box)", quantity_display)
	row_top -= rh

	# -----------------------------
	# Barcode row (full width cell)
	# -----------------------------
	rh = row_heights[3]
	c.rect(label_x, row_top - rh, label_w, rh)

	# Reserve space for human-readable text at the bottom
	text_size = 12
	text_pad_bottom = 16
	text_area_h = text_size + 6
	top_pad = 16
	bottom_pad = text_area_h + text_pad_bottom

	size_height_target = rh - top_pad - bottom_pad
	if size_height_target < 30:
		size_height_target = 30

	size_width_target = label_w - 60

	# Initial barcode baseline
	bar_scale = 100
	initial_bar_width = ((1.2 / 100) * bar_scale)
	initial_bar_height = ((40.0 / 100) * bar_scale)
	initial_barcode = code128.Code128(barcode_data, barHeight=initial_bar_height, barWidth=initial_bar_width)

	size_width_init = initial_barcode.width
	size_height_init = initial_barcode.height

	# Width-based sizing formula
	size_diff_percent = (size_width_init / 100.0) * size_width_target

	final_width = size_diff_percent
	final_height = (size_height_init / 100.0) * size_diff_percent

	# Fit clamps
	if final_width > size_width_target:
		scale_w = size_width_target / final_width
		final_width *= scale_w
		final_height *= scale_w

	if final_height > size_height_target:
		scale_h = size_height_target / final_height
		final_width *= scale_h
		final_height *= scale_h

	# Convert target total width into barWidth
	bar_width_final = initial_bar_width * (final_width / size_width_init)
	if bar_width_final < 0.2:
		bar_width_final = 0.2

	barcode_obj = code128.Code128(barcode_data, barHeight=final_height, barWidth=bar_width_final)

	barcode_w = barcode_obj.width
	barcode_h = barcode_obj.height
	barcode_x = label_x + (label_w - barcode_w) / 2
	barcode_y = (row_top - rh) + top_pad + (size_height_target - barcode_h) / 2

	# Intentional offset to prevent overlap
	barcode_obj.drawOn(c, barcode_x, (barcode_y + 15))

	# Human-readable text at the bottom (same as barcode payload)
	c.setFont(FONT_REGULAR, text_size)
	c.drawCentredString(label_x + label_w / 2, (row_top - rh) + text_pad_bottom, barcode_data)

	row_top -= rh

	# -----------------------------
	# Order Index
	# -----------------------------
	rh = row_heights[4]
	draw_row(row_top, rh, "Order Index", "(Nth item in your order)", order_index_display, shrinkable=True)
	row_top -= rh

	# -----------------------------
	# Order No. (shrink to fit)
	# -----------------------------
	rh = row_heights[5]
	draw_row(row_top, rh, "Order No.", "(Your order number)", order_display, shrinkable=True)
	row_top -= rh

	# -----------------------------
	# Box No.
	# -----------------------------
	rh = row_heights[6]
	draw_row(row_top, rh, "Box No.", "(Nth box in this order)", str(carton_no))
	row_top -= rh

	# -----------------------------
	# Weight
	# -----------------------------
	rh = row_heights[7]
	draw_row(row_top, rh, "Weight", "(Box weight)", weight_display, shrinkable=True)
	row_top -= rh




def build_output_pdf_filename(csv_path, data_rows):
	base_name = os.path.splitext(os.path.basename(csv_path))[0]
	base_name = safe_filename_component(base_name, default="labels")

	if not data_rows:
		return f"{base_name}.pdf"

	company_values = []
	order_values = []

	for row in data_rows:
		company_value = truncate_component(safe_filename_component(row.get("company_raw"), default="Company"), 80)
		order_value = truncate_component(safe_filename_component(row.get("order_raw"), default="OrderNo"), 80)

		if company_value and company_value not in company_values:
			company_values.append(company_value)
		if order_value and order_value not in order_values:
			order_values.append(order_value)

	company_part = company_values[0] if company_values else base_name
	order_part = order_values[0] if order_values else "OrderNo"

	return f"{company_part} - Order {order_part}.pdf"



def generate_separate_pdf(company_raw, artikel_raw, qty_raw, order_index_raw, order_raw, carton_no, weight_raw):
	company = sanitize_any(company_raw, default="COMPANY_NAME")
	artikel_display = sanitize_any(artikel_raw, default="")
	order_display = sanitize_any(order_raw, default="")
	order_index_display = sanitize_any(order_index_raw, default="1")

	company_safe = truncate_component(safe_filename_component(company, default="Company"), 40)
	order_safe = truncate_component(safe_filename_component(order_display, default="NoOrderNo"), 40)
	order_index_safe = truncate_component(safe_filename_component(order_index_display, default="OrderIndex"), 20)

	pdf_file = f"{company_safe} - Order {order_safe}, Index {order_index_safe}, Article {artikel_display}, Box {carton_no}.pdf"
	c = canvas.Canvas(pdf_file, pagesize=A4)
	draw_label_page(c, company_raw, artikel_raw, qty_raw, order_index_raw, order_raw, carton_no, weight_raw)
	c.save()
	print(f"PDF saved as {pdf_file}")



def generate_multipage_pdf(pdf_file, pages):
	c = canvas.Canvas(pdf_file, pagesize=A4)
	for page in pages:
		draw_label_page(
			c,
			page["company_raw"],
			page["artikel_raw"],
			page["qty_raw"],
			page["order_index_raw"],
			page["order_raw"],
			page["carton_no"],
			page["weight_raw"],
		)
		c.showPage()
	c.save()
	print(f"PDF saved as {pdf_file}")

# -----------------------------
# CSV mode (only)
# -----------------------------
def run_csv(csv_path):
	"""CSV format: semicolon-separated ; values quoted with "

	Supports header-based matching.

	Expected columns (names can vary; matching is best-effort):
	- COMPANY (new): displayed in the header row
	- Artikel-Nr / Article No / Item No / SKU (required)
	- Quantity / Stueckzahl / Quantity / Qty (required)
	- Order Index / Batch / Chargen-Nr (required)
	- Bestell-Nr / Order (optional)
	- Karton-Nr / Carton (required; if value > 1, generates that many labels with carton numbers 1..N)
	- Weight / Gewicht (optional)
	"""
	with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
		reader = csv.reader(f, delimiter=";", quotechar='"')

		header = None
		colmap = None
		data_rows = []

		for row in reader:
			# Skip empty rows
			if not row or all((cell is None or str(cell).strip() == "") for cell in row):
				continue

			# First non-empty row is header
			if header is None:
				header = row
				colmap = build_column_map(header)
				continue

			company_raw = get_cell(row, colmap.get("company"), default="COMPANY_NAME")
			artikel_raw = get_cell(row, colmap.get("article"), default="")
			qty_raw = get_cell(row, colmap.get("quantity"), default="0")
			order_index_raw = get_cell(row, colmap.get("order_index"), default="1")
			order_raw = get_cell(row, colmap.get("order"), default="")
			karton_raw = get_cell(row, colmap.get("carton"), default="1")
			weight_raw = get_cell(row, colmap.get("weight"), default="")

			karton_digits = sanitize_numeric_keep_zeros(karton_raw, default="1")
			try:
				k_total = int(karton_digits) if karton_digits else 1
			except ValueError:
				k_total = 1

			if k_total < 1:
				k_total = 1

			data_rows.append({
				"company_raw": company_raw,
				"artikel_raw": artikel_raw,
				"qty_raw": qty_raw,
				"order_index_raw": order_index_raw,
				"order_raw": order_raw,
				"weight_raw": weight_raw,
				"k_total": k_total,
			})

	mode = str(CARTON_NUMBERING_MODE).strip().lower()
	valid_modes = {"per_article", "per_order"}
	if mode not in valid_modes:
		print(f'[WARNING] Invalid CARTON_NUMBERING_MODE: {CARTON_NUMBERING_MODE}. Falling back to "per_order".')
		mode = "per_order"

	output_mode = str(PDF_OUTPUT_MODE).strip().lower()
	valid_output_modes = {"multipage", "separate"}
	if output_mode not in valid_output_modes:
		print(f'[WARNING] Invalid PDF_OUTPUT_MODE: {PDF_OUTPUT_MODE}. Falling back to "multipage".')
		output_mode = "multipage"

	order_counters = {}
	pages_to_generate = []

	for row_data in data_rows:
		company_raw = row_data["company_raw"]
		artikel_raw = row_data["artikel_raw"]
		qty_raw = row_data["qty_raw"]
		order_index_raw = row_data["order_index_raw"]
		order_raw = row_data["order_raw"]
		weight_raw = row_data["weight_raw"]
		k_total = row_data["k_total"]

		order_key = sanitize_any(order_raw, default="")

		for _ in range(k_total):
			if mode == "per_order" and order_key:
				current_carton_no = order_counters.get(order_key, 0) + 1
				order_counters[order_key] = current_carton_no
			else:
				current_carton_no = _ + 1

			page_data = {
				"company_raw": company_raw,
				"artikel_raw": artikel_raw,
				"qty_raw": qty_raw,
				"order_index_raw": order_index_raw,
				"order_raw": order_raw,
				"carton_no": str(current_carton_no),
				"weight_raw": weight_raw,
			}

			if output_mode == "multipage":
				pages_to_generate.append(page_data)
			else:
				generate_separate_pdf(**page_data)

	if output_mode == "multipage":
		pdf_file = build_output_pdf_filename(csv_path, data_rows)
		generate_multipage_pdf(pdf_file, pages_to_generate)

# -----------------------------
# Entry
# -----------------------------
if __name__ == "__main__":
	if len(sys.argv) < 2:
		print("[ERROR] No CSV file provided.")
		print("Usage: python pyStickerCreator.py <path-to-csv>")
		sys.exit(1)

	csv_arg = sys.argv[1]
	if not os.path.isfile(csv_arg):
		print(f"[ERROR] CSV file not found: {csv_arg}")
		sys.exit(1)

	run_csv(csv_arg)
