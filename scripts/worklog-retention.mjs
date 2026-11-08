import admin from "firebase-admin";
import ExcelJS from "exceljs";
import crypto from "node:crypto";
import fs from "node:fs/promises";

const required = ["FIREBASE_SERVICE_ACCOUNT_JSON", "CLOUDINARY_CLOUD_NAME", "CLOUDINARY_API_KEY", "CLOUDINARY_API_SECRET"];
for (const key of required) if (!process.env[key]) throw new Error(`Missing required environment variable: ${key}`);

const enabledDelete = process.env.ENABLE_PERMANENT_DELETE === "true";
const projectId = process.env.FIREBASE_PROJECT_ID || "proto-beta1";
const cloud = process.env.CLOUDINARY_CLOUD_NAME;
const cutoff = Math.floor(Date.now() / 1000) - 180 * 24 * 60 * 60;
const timestamp = Math.floor(Date.now() / 1000);
const runId = new Date().toISOString().replace(/[:.]/g, "-");
const report = { runId, cutoff: new Date(cutoff * 1000).toISOString(), candidates: 0, backedUp: 0, deletedDocuments: 0, deletedMedia: 0, skipped: [], errors: [], deletionEnabled: enabledDelete };

const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT_JSON);
admin.initializeApp({ credential: admin.credential.cert(serviceAccount), projectId });
const db = admin.firestore();
const collection = db.collection("artifacts").doc("shift-handover-app").collection("public").doc("data").collection("handovers");

const toSeconds = value => {
  if (value && typeof value.toMillis === "function") return Math.floor(value.toMillis() / 1000);
  if (value && typeof value.seconds === "number") return value.seconds;
  if (typeof value === "number") return value;
  return 0;
};
const safeCell = value => value == null ? "" : String(value);
const sign = params => {
  const source = Object.keys(params).sort().map(k => `${k}=${params[k]}`).join("&") + process.env.CLOUDINARY_API_SECRET;
  return crypto.createHash("sha1").update(source).digest("hex");
};

async function cloudinaryUpload(buffer, publicId) {
  const params = { folder: "worklog-backups", public_id: publicId, timestamp: String(timestamp) };
  const form = new FormData();
  form.append("file", new Blob([buffer]), `${publicId}.xlsx`);
  for (const [key, value] of Object.entries(params)) form.append(key, value);
  form.append("api_key", process.env.CLOUDINARY_API_KEY);
  form.append("signature", sign(params));
  const response = await fetch(`https://api.cloudinary.com/v1_1/${cloud}/raw/upload`, { method: "POST", body: form });
  const body = await response.json();
  if (!response.ok) throw new Error(`Cloudinary backup upload failed: ${body.error?.message || response.status}`);
  return body;
}

async function destroyMedia(item) {
  if (!item?.publicId) throw new Error("Attachment is missing publicId; refusing unsafe media deletion.");
  const resourceType = item.type === "video" ? "video" : "image";
  const params = { public_id: item.publicId, timestamp: String(Math.floor(Date.now() / 1000)) };
  const form = new FormData();
  form.append("public_id", params.public_id);
  form.append("timestamp", params.timestamp);
  form.append("api_key", process.env.CLOUDINARY_API_KEY);
  form.append("signature", sign(params));
  const response = await fetch(`https://api.cloudinary.com/v1_1/${cloud}/${resourceType}/destroy`, { method: "POST", body: form });
  const body = await response.json();
  if (!response.ok || !["ok", "not found"].includes(body.result)) {
    throw new Error(`Cloudinary media deletion failed for ${item.publicId}: ${body.result || body.error?.message || response.status}`);
  }
  return body.result === "ok";
}

const snapshot = await collection.get();
const candidates = snapshot.docs.map(doc => ({ ref: doc.ref, id: doc.id, data: doc.data() }))
  .filter(x => x.data.completed === true && toSeconds(x.data.createdAt) > 0 && toSeconds(x.data.createdAt) <= cutoff)
  .sort((a, b) => toSeconds(a.data.createdAt) - toSeconds(b.data.createdAt));
report.candidates = candidates.length;
console.log(`Found ${candidates.length} completed records older than 180 days.`);

if (candidates.length === 0) {
  await fs.writeFile("worklog-retention-report.json", JSON.stringify(report, null, 2));
  console.log("No eligible records. No backup or deletion required.");
  process.exit(0);
}

const workbook = new ExcelJS.Workbook();
workbook.creator = "WorkLog automated archive";
workbook.created = new Date();
const sheet = workbook.addWorksheet("조치완료 백업");
sheet.columns = [
  { header: "Firestore 문서 ID", key: "id", width: 30 },
  { header: "작성일", key: "createdAt", width: 22 },
  { header: "라인", key: "line", width: 14 },
  { header: "근무조", key: "shift", width: 14 },
  { header: "작성자", key: "author", width: 20 },
  { header: "내용", key: "memo", width: 60 },
  { header: "조치완료", key: "completed", width: 12 },
  { header: "첨부 메타데이터(백업 후 원본 삭제 대상)", key: "attachments", width: 70 }
];
sheet.getRow(1).font = { bold: true };
for (const item of candidates) {
  const d = item.data;
  const attachments = Array.isArray(d.attachments) ? d.attachments : [];
  sheet.addRow({
    id: item.id,
    createdAt: toSeconds(d.createdAt) ? new Date(toSeconds(d.createdAt) * 1000).toISOString() : "",
    line: safeCell(d.line || "미분류"),
    shift: safeCell(d.shift),
    author: safeCell(d.author),
    memo: safeCell(d.memo),
    completed: "완료",
    attachments: attachments.map(a => `${a.type || "image"} | ${a.name || ""} | ${a.publicId || ""} | ${a.url || ""}`).join("\n")
  });
}
sheet.views = [{ state: "frozen", ySplit: 1 }];
const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
if (sheet.rowCount !== candidates.length + 1) throw new Error("Backup row-count verification failed.");
const publicId = `worklog-quarterly-${runId}`;
const uploaded = await cloudinaryUpload(buffer, publicId);
if (!uploaded.secure_url || uploaded.resource_type !== "raw" || Number(uploaded.bytes) !== buffer.length) {
  throw new Error("Cloudinary upload verification failed; deletion is blocked.");
}
report.backedUp = candidates.length;
report.backup = { publicId: uploaded.public_id, secureUrl: uploaded.secure_url, bytes: uploaded.bytes };
console.log(`Verified Excel backup: ${uploaded.secure_url} (${candidates.length} rows)`);

if (enabledDelete) {
  for (const item of candidates) {
    try {
      const attachments = Array.isArray(item.data.attachments) ? item.data.attachments : [];
      for (const attachment of attachments) {
        const deleted = await destroyMedia(attachment);
        if (deleted) report.deletedMedia++;
      }
      await item.ref.delete();
      report.deletedDocuments++;
    } catch (error) {
      report.errors.push({ id: item.id, error: String(error) });
      console.error(`Deletion stopped for ${item.id}: ${error.message}`);
    }
  }
} else {
  report.skipped.push("Permanent deletion is disabled. Backup completed; all Firestore records and media were retained.");
}
await fs.writeFile("worklog-retention-report.json", JSON.stringify(report, null, 2));
console.log(JSON.stringify(report, null, 2));
