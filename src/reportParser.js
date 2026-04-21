// Parses a Report 1145 Excel file and maps it to the 23-column Tabelle1 format
// used by the FUTURELOG XML generator.
//
// VBA reference (Get_Data_From_File1145):
//   Source data starts at row 5 of the 1145 file.
//   B          -> A (sequence will be overwritten as pos = row-1)
//   B:N        -> B:N   (type, DE, FR, IT, GB, item, artNo, EAN, MfgNum, Producer, OU, CU, CU/OU)
//   Q:R        -> Q:R   (Origin, Customs no.)
//   T:V        -> T:V   (availability placeholder, offer start, offer end)
//   O:P        -> X:Y   (Price/OU, Scaled price)  [scratch cols]
//   S          -> Z     (lead time)               [scratch col]
//
// Post-processing:
//   - Round col X to 2 decimals
//   - Col O = IF(Y is empty / 0 / not numeric, X, Y)   (scaled-price preference)
//   - Col S (availability) = IF(O == Y, Z, 0)          (only use lead time when scaled price used)
//
// Final 23 cols of Tabelle1 we care about for XML:
//   1 pos, 2 type, 3 DE, 4 FR, 5 IT, 6 GB, 7 Item, 8 ItemNo, 9 EAN,
//   10 ManArtId, 11 ManLiefID, 12 OU, 13 CU, 14 CU/OU, 15 Price/OU,
//   16 Price Level, 17 Origin, 18 customs no., 19 availability,
//   20 specification URL, 21 offer start, 22 offer end, 23 Customer ID

const HEADER_ROW_INDEX = 3; // 0-based: row 4 in Excel is the header row with "ID, Type, ..."

// Sentinel used internally so the validator can distinguish "cell was #N/A" from "cell was blank".
// The XML generator treats this as empty.
export const NA_MARKER = "__NA__";

function cleanCell(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") {
    const t = v.trim();
    if (t === "#N/A" || t === "N/A" || t === "#N/A N/A") return NA_MARKER;
    return t;
  }
  if (typeof v === "number") {
    if (Number.isNaN(v)) return "";
    return v;
  }
  if (v instanceof Date) {
    // dd.mm.yyyy -> will be stripped of dots later for OFFSTART/OFFEND
    const d = String(v.getDate()).padStart(2, "0");
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const y = v.getFullYear();
    return `${d}.${m}.${y}`;
  }
  return String(v);
}

function toNumericOrEmpty(v) {
  if (v === "" || v === null || v === undefined) return "";
  if (typeof v === "number") return v;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "") return "";
  const n = Number(s);
  return Number.isFinite(n) ? n : "";
}

/**
 * @param {Array<Array<any>>} aoa  sheet parsed as array-of-arrays (header:1 from SheetJS)
 * @returns {Array<Object>}        rows in Tabelle1 format
 */
export function parseReport1145(aoa) {
  if (!Array.isArray(aoa) || aoa.length <= HEADER_ROW_INDEX) {
    throw new Error("Report 1145 appears to be empty or malformed (no data below header row 4).");
  }

  // Verify header row shape — help the user if they picked the wrong file
  const headerRow = aoa[HEADER_ROW_INDEX] || [];
  const looksValid =
    String(headerRow[0] ?? "").toUpperCase() === "ID" &&
    String(headerRow[1] ?? "").toLowerCase() === "type" &&
    String(headerRow[7] ?? "").toLowerCase().startsWith("article no");
  if (!looksValid) {
    throw new Error(
      "This doesn't look like a Report 1145 file. Expected headers 'ID', 'Type', 'Article no.' on row 4."
    );
  }

  const rows = [];

  // Source rows start at Excel row 5 => zero-based index 4
  for (let r = HEADER_ROW_INDEX + 1; r < aoa.length; r++) {
    const src = aoa[r] || [];
    // Stop at the first empty "Type" column — matches the VBA `Do Until Cells(Z, 2) = ""`
    const typeCol = cleanCell(src[1]); // B in source
    if (typeCol === "" || typeCol === 0) break;

    const priceOU = toNumericOrEmpty(src[14]); // O  unit price
    const scaled = toNumericOrEmpty(src[15]); // P  scaled price
    const leadTime = cleanCell(src[18]); // S  lead time

    // VBA: round X (scaled) to 2 decimals, then O = (Y empty/0/non-numeric ? X : Y)
    const scaledRounded =
      typeof scaled === "number" ? Math.round(scaled * 100) / 100 : "";

    let finalPrice;
    let usedScaled = false;
    if (scaledRounded === "" || scaledRounded === 0) {
      finalPrice = priceOU === "" ? 0 : priceOU;
    } else {
      finalPrice = scaledRounded;
      usedScaled = true;
    }

    // availability: if final price == scaled price (i.e. scaled branch) => lead time else "0"
    const availability = usedScaled ? leadTime : "0";

    // Build the 23-col Tabelle1 row
    const row = {
      pos: rows.length + 1,
      recordType: cleanCell(src[1]), // B  -> type
      descDE: cleanCell(src[2]), // C
      descFR: cleanCell(src[3]), // D
      descIT: cleanCell(src[4]), // E
      descGB: cleanCell(src[5]), // F  (English)
      descExtra: cleanCell(src[6]), // G  second language / generic item name
      itemNo: String(cleanCell(src[7])).trim(), // H  Article no
      ean: String(cleanCell(src[8])).trim(), // I  GTIN
      manArtId: cleanCell(src[9]), // J
      manLiefID: cleanCell(src[10]), // K
      ou: cleanCell(src[11]), // L  OU
      cu: cleanCell(src[12]), // M  CU
      cuou: cleanCell(src[13]), // N  CU/OU
      priceOU: finalPrice, // O  final unit price
      priceLevel: "", // P  (blank in 1145 import)
      origin: cleanCell(src[16]), // Q  country of origin
      customsNo: cleanCell(src[17]), // R
      availability: availability, // S
      specUrl: cleanCell(src[19]), // T  item link supplier
      offerStart: cleanCell(src[20]), // U
      offerEnd: cleanCell(src[21]), // V
      customerId: "0000", // W  defaulted per user requirement
    };
    rows.push(row);
  }

  return rows;
}
