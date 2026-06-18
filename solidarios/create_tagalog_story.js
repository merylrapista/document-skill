const docx = require("docx");
const fs = require("fs");

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  PageOrientation,
  convertInchesToTwip,
} = docx;

const storyTitle = "Ang Paglalakbay ni Lola Ising";
const author = "Isinulat ni: AI Katutubo";
const date = "Petsa: Hunyo 18, 2026";

// Tagalog short story content (a heartfelt story about a grandmother's journey)
const storyParts = [
  {
    type: "heading",
    text: storyTitle,
  },
  {
    type: "subheading",
    text: author,
  },
  {
    type: "subheading",
    text: date,
  },
  {
    type: "paragraph",
    text: "Sa isang maliit na baryo sa probinsya ng Quezon, namumuhay si Lola Ising kasama ang kanyang tatlong apo. Sa kabila ng kanyang katandaan, nananatiling matalas ang kanyong isip at malakas ang kanyang loob. Bawat umaga, gumigising siya ng maaga upang magluto ng kanilang almusal at maghanda sa mga gawain sa araw.",
  },
  {
    type: "paragraph",
    text: "\"Lola, pahingi po ng tubig,\" ani ng bunso niyang apo na si Marco. Ngumiti si Lola Ising at dali-daling naghugas ng baso. Mahilig niyang bigyan ang kanyang mga apo ng lahat ng kanilang nais, sapagkat siya ang nagsilbing pangalawang ina sa kanila matapos pumanaw ang kanilang mga magulang sa isang aksidente noong nakaraang taon.",
  },
  {
    type: "paragraph",
    text: "Isang araw, dumating ang sulat mula sa Maynila. Nag-aanyaya ang kanyang anak na si Rosa na bumisita sa kanila sa lungsod. Masaya si Lola Ising ngunit may pangamba rin sa kanyang puso. Matagal na niyang hindi nakikita ang kanyang anak, at natatakot siya sa magiging pagbabago sa kanyang mga apo kapag sila ay nasa lungsod na.",
  },
  {
    type: "paragraph",
    text: "Sa gabing iyon, habang natutulog ang mga bata, nakaupo si Lola Ising sa veranda at tinitignan ang mga bituin. Naisip niya ang kanyang asawa na siyang namayapa na noong nakaraang dekada. \"Sana'y nandito ka para tulungan akong magpasya, Ka Gorio,\" bulong niya sa hangin.",
  },
  {
    type: "paragraph",
    text: "Kinaumagahan, nag-usap sina Lola Ising at ang kanyang mga apo. Ipinaliwanag niya na mayroon silang mahalagang desisyon na dapat gawin. \"Mga apo, gusto ko sanang malaman ninyo kung ano ang nararapat nating gawin,\" ani Lola Ising. \"Kung tayo ay pupunta sa Maynila, magkakaroon kayo ng magandang kinabukasan. Ngunit kung mananatili tayo rito, makakapagsimula tayo ng bagong bukas sa ating sariling lupa.\"",
  },
  {
    type: "paragraph",
    text: "Matapos ang mahabang pag-iisip, nagpasya silang manatili sa baryo. Nais ni Lola Ising na bigyan ang kanyang mga apo ng simpleng ngunit tunay na buhay - isang buhay na puno ng pagmamahalan, pagtutulungan, at paggalang sa kalikasan. \"Tayo ay magtatayo ng sarili nating pamana rito,\" ani niya. \"Ang lupaing ito ay magiging yaman natin, hindi lang sa materyal na paraan, kundi pati na rin sa espirituwal.\"",
  },
  {
    type: "paragraph",
    text: "At sa gayon, sa pangunguna ni Lola Ising, nagsimula silang magtayo ng kanilang sariling hardin, nagturo ng mga bata tungkol sa pagsasaka, at nagbahagi ng kanilang ani sa mga kapitbahay. Hindi naging madali ang buhay, ngunit naging masaya at makabuluhan ito dahil sa kanilang pagkakaisa at pagmamahal sa isa't isa.",
  },
  {
    type: "paragraph",
    text: "Iyon ay simula pa lamang. Sa mga susunod na taon, naging tanyag ang baryo nila dahil sa kanilang mga produkto at sa pagtuturo ng mga bata. Naging inspirasyon si Lola Ising sa maraming tao - isang matandang babae na nagpakita na ang tunay na yaman ay hindi nasusukat sa salapi, kundi sa mga bagay na walang presyo - pagmamahal, pagtitiwala, at pagkakaisa ng pamilya.",
  },
  {
    type: "ending",
    text: "— Wakas —",
  },
];

const children = [];

storyParts.forEach((part) => {
  if (part.type === "heading") {
    children.push(
      new Paragraph({
        text: part.text,
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
      })
    );
  } else if (part.type === "subheading") {
    children.push(
      new Paragraph({
        text: part.text,
        alignment: AlignmentType.CENTER,
        italics: true,
        spacing: { after: 100 },
      })
    );
  } else if (part.type === "paragraph") {
    children.push(
      new Paragraph({
        text: part.text,
        alignment: AlignmentType.JUSTIFY,
        spacing: { line: 360, after: 200 },
        children: [
          new TextRun({
            text: part.text,
            size: 24,
            font: "Calibri",
          }),
        ],
      })
    );
  } else if (part.type === "ending") {
    children.push(
      new Paragraph({
        text: part.text,
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 200 },
        children: [
          new TextRun({
            text: part.text,
            italics: true,
            size: 24,
          }),
        ],
      })
    );
  }
});

const doc = new Document({
  creator: "AI Katutubo",
  title: storyTitle,
  description: "Isang maikling kwento tungkol sa paglalakbay ni Lola Ising",
  styles: {
    default: {
      document: {
        run: {
          font: "Calibri",
          size: 24,
        },
      },
    },
  },
  sections: [
    {
      properties: {
        page: {
          margin: {
            top: convertInchesToTwip(1.25),
            right: convertInchesToTwip(1.25),
            bottom: convertInchesToTwip(1.25),
            left: convertInchesToTwip(1.25),
          },
          size: {
            orientation: PageOrientation.PORTRAIT,
          },
        },
      },
      children: children,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("tagalog_short_story.docx", buffer);
  console.log("Tagalog short story saved as tagalog_short_story.docx");
});
