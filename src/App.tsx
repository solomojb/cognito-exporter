import { useRef, useState } from "react";
import "./App.css";
import * as xlsx from "xlsx";
import { Document, ISectionOptions, Packer, Paragraph, TextRun, ColumnBreak } from "docx";
import saveAs from "file-saver";

const convertField = (field: Record<string, string>) => {
    return Object.entries(field)
        .filter(([key, value]) => {
            if (Number.isInteger(value)) {
                const num = parseInt(value);
                if (!num) {
                    return false;
                }
            }
            if (
                key.includes("Entry_") ||
                key.includes("5050Raffle_Total") ||
                key.includes("Name_First") ||
                key.includes("Name_Last") ||
                key.includes("Email") ||
                key.includes("Phone") ||
                key.includes("TotalCost") ||
                key.includes("HowDoYouPlanOnPaying")

            ) {
                return false;
            }
            return true;
        })
        .map(([key, value]) => {
            let label = key.replace("_", " ");
            if (key.includes("_Id")) {
                label = "Id";
            } else if (key.includes("_Tickets")) {
                label = key.substring(0, key.indexOf("_Tickets"));
            }
            label = label.replace("_", " ").trim();
            return new TextRun({
                text: label ? `${label}: ${value}` : ` ${value}`,
                break: 2,
                size: '14pt',
            });
        });
}

export const App = () => {
    const inputRef = useRef<HTMLInputElement>(null);
    const [fields, setFields] = useState<
        Record<string, string>[] | undefined
    >();
    const handleFileInput = (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        const reader = new FileReader();
        if (!file) {
            return;
        }
        reader.readAsArrayBuffer(file);
        reader.onload = () => {
            const buffer = reader.result as ArrayBuffer;

            if (!buffer) {
                return;
            }
            const workbook = xlsx.read(buffer, { type: "buffer" });
            const sheet = Object.values(workbook.Sheets)[0];
            const json =
                xlsx.utils.sheet_to_json<Record<string, string>>(sheet);
            console.log(Object.keys(json[0]));
            setFields(json);
        };
    };

    const convert = () => {
        if (!fields) {
            return;
        }
        const sections: ISectionOptions[] = [];
        for (let i = 0; i < fields.length; i += 2) {
            const first = fields[i];
            const field1 = convertField(first);
            const second = fields[i + 1];
            const field2 = convertField(second);
            const page: ISectionOptions = {
                properties: {
                    column: {
                        space: 708,
                        equalWidth: true,
                        count: 2
                    }
                },
                children: [
                    new Paragraph({
                        children: [...field1, new ColumnBreak(), ...field2],
                    }),
                ],
            };
            sections.push(page);

        }
        const document = new Document({
            sections,
        });
        Packer.toBlob(document).then((blob) => {
            saveAs(blob, `foo.docx`);
        });
    };

    return (
        <>
            <button
                onClick={() => {
                    inputRef.current?.click();
                }}
            >
                Browse
            </button>

            <input
                hidden
                ref={inputRef}
                type="file"
                onChange={handleFileInput}
            />
            {fields && <button onClick={convert}>convert</button>}
        </>
    );
};
