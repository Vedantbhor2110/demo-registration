let timer;

document.addEventListener("input", () => {
    clearTimeout(timer);
    timer = setTimeout(saveToWord, 1200);
});

function saveToWord() {
    const name = document.getElementById("name")?.value || "";
    const email = document.getElementById("email")?.value || "";
    const password = document.getElementById("password")?.value || "";

    if (!name && !email && !password) return;

    const { Document, Packer, Paragraph, TextRun } = window.docx;

    const doc = new Document({
        sections: [{
            children: [

                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Amazon Demo Registration",
                            bold: true,
                            size: 28
                        })
                    ]
                }),

                new Paragraph({
                    children: [ new TextRun("Name: " + name) ]
                }),

                new Paragraph({
                    children: [ new TextRun("Email: " + email) ]
                }),

                new Paragraph({
                    children: [ new TextRun("Password: " + password) ]
                })
            ]
        }]
    });

    Packer.toBlob(doc).then(blob => {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "amazon_demo_data.docx";
        link.click();
    });
}
