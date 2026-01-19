const fs = require("fs");
const path = require("path");

const winPath = path.win32;
const folder = '.';

// Удаляем недопустимые символы Windows
function sanitize(str) {
    return str.replace(/[<>:"/\\|?*]/g, "_");
}

// Декодирование MIME encoded-word (=?UTF-8?Q?...?=)
function decodeMimeWord(str) {
    if (!str) return str;

    return str.replace(/=\?([^?]+)\?([QBqb])\?([^?]+)\?=/g, (_, charset, encoding, text) => {
        charset = charset.toLowerCase();
        encoding = encoding.toUpperCase();

        let buf;

        if (encoding === "Q") {
            // Quoted-printable → сырые байты
            const qp = text
                .replace(/_/g, " ")
                .replace(/=([A-Fa-f0-9]{2})/g, (_, hex) =>
                    String.fromCharCode(parseInt(hex, 16))
                );

            buf = Buffer.from(qp, "binary");
        }

        if (encoding === "B") {
            buf = Buffer.from(text, "base64");
        }

        // Декодируем строго как UTF‑8
        if (charset === "utf-8" || charset === "utf8") {
            return buf.toString("utf8");
        }

        // fallback
        return buf.toString();
    });
}

// Извлекаем email из строки
function cleanEmail(str) {
    if (!str) return "unknown";

    str = decodeMimeWord(str);

    const match = str.match(/<([^>]+)>/);
    const email = match ? match[1] : str.trim();

    return sanitize(email);
}

// Форматирование даты
function formatDate(date) {
    const pad = n => String(n).padStart(2, "0");
    return (
        date.getFullYear() + "-" +
        pad(date.getMonth() + 1) + "-" +
        pad(date.getDate()) + "_" +
        pad(date.getHours()) + "-" +
        pad(date.getMinutes()) + "-" +
        pad(date.getSeconds())
    );
}

fs.readdirSync(folder).forEach(file => {
    if (!file.toLowerCase().endsWith(".eml")) return;

    const fullPath = winPath.join(folder, file);

    let content = fs.readFileSync(fullPath, "utf8");
    content = content.replace(/^\uFEFF/, ""); // убираем BOM

    const dateHeader = content.match(/^Date:\s*(.+)$/mi);
    const fromHeader = content.match(/^From:\s*(.+)$/mi);
    const toHeader = content.match(/^To:\s*(.+)$/mi);

    const dateStr = dateHeader ? dateHeader[1] : null;
    const fromStr = fromHeader ? fromHeader[1] : null;
    const toStr = toHeader ? toHeader[1] : null;

    let date;
    try {
        date = dateStr ? new Date(dateStr) : null;
        if (!date || isNaN(date.getTime())) throw new Error();
    } catch {
        console.log(`Не удалось разобрать дату: ${file}`);
        return;
    }

    const formattedDate = formatDate(date);
    const fromEmail = cleanEmail(fromStr);
    const toEmail = cleanEmail(toStr);

    const newName = `${formattedDate}__${fromEmail}__${toEmail}.eml`;
    // const newName = `${formattedDate}__${fromEmail}__${toEmail}.eml`;
    const newPath = winPath.join(folder, newName);

    fs.renameSync(fullPath, newPath);
    console.log(`Переименовано: ${file} → ${newName}`);
});
