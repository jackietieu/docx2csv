import Papa from "papaparse";
import "./App.css";
import mammoth from "mammoth";

function App() {
  const handleFileUpload = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];

    if (file) {
      const text = await extractTextFromDocx(file);
      const data = parseText(text);
      const csv = Papa.unparse(data);
      const blob = new Blob([csv], { type: "text/csv" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = file.name.replace(".docx", ".csv");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }
  };

  return (
    <>
      <div
        style={{
          display: "flex",
          gap: "8px",
          flexDirection: "column",
          justifyContent: "center",
          alignItems: "center",
        }}
      >
        <label htmlFor="file-input">.docx only</label>
        <input
          id="file-input"
          type="file"
          accept=".docx"
          onChange={handleFileUpload}
          className="hidden"
        />
        <span>
          Just upload the DOCX file and it'll automatically download the CSV
          file for you
        </span>
      </div>
    </>
  );
}

const headings = [
  "Abstract Number",
  "Subheading",
  "Title",
  "First Author Name",
  "First Author Gender",
  "First Author Affiliation",
  "Last Author Name",
  "Last Author Gender",
  "Last Author Affiliation",
];

async function extractTextFromDocx(file: File) {
  const result = await mammoth.extractRawText({
    arrayBuffer: await file.arrayBuffer(),
  });
  return result.value;
}

// Function to parse the extracted text
function parseText(text: string) {
  const lines = text
    .split("\n")
    .filter((line) => line !== "\n")
    .map((line) => line.replace("\n", ""));
  const data: Record<string, string>[] = [];
  let entryData: Record<string, string> = {};

  lines.forEach((line) => {
    if (!!line.match(/^(Abstract )?#\d+/)) {
      entryData["Abstract Number"] = line
        .replace("Abstract ", "")
        .replace("#", "")
        .trim();
    } else if (line.startsWith("Subheading:")) {
      entryData["Subheading"] = line.replace("Subheading:", "").trim();
    } else if (line.startsWith("Title:")) {
      entryData["Title"] = line.replace("Title:", "").trim();
    } else if (line.startsWith("First Author:")) {
      const match = line.match(/First Author: (.*) \((male|female)\).*,(.*)/);
      if (match) {
        entryData["First Author Name"] = match[1].trim();
        entryData["First Author Gender"] = match[2];
        entryData["First Author Affiliation"] = match[3];
      }
    } else if (line.startsWith("Last Author:")) {
      const match = line.match(/Last Author: (.*) \((male|female)\).*,(.*)/);
      if (match) {
        entryData["Last Author Name"] = match[1].trim();
        entryData["Last Author Gender"] = match[2];
        entryData["Last Author Affiliation"] = match[3];
      }
    }

    if (headings.every((heading) => entryData[heading])) {
      data.push(entryData);

      // reset and start next entry
      entryData = {};
    }
  });

  return data;
}

export default App;
