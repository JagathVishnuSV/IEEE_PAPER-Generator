document.getElementById("paperForm").addEventListener("submit", async function (e) {
    e.preventDefault();
  
    const form = e.target;
    const data = {
      title: form.title.value,
      authors: form.authors.value.split(","),
      affiliations: form.affiliations.value.split(","),
      emails: form.emails.value.split(","),
      abstract: form.abstract.value,
      keywords: form.keywords.value.split(","),
      sections: [],
      references: form.references.value.split("\n")
    };
  
    const sectionElements = document.querySelectorAll(".section");
    sectionElements.forEach(section => {
      const heading = section.querySelector(".heading").value;
      const content = section.querySelector(".content").value;
  
      data.sections.push({
        heading,
        content,
        images: [],
        tables: [],
        formulas: []
      });
    });
  
    try {
      const response = await fetch("http://localhost:8000/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data)
      });
  
      if (!response.ok) throw new Error("Failed to generate paper.");
  
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "ieee_paper.docx";
      a.click();
    } catch (error) {
      alert("Error: " + error.message);
    }
  });
  
  function addSection() {
    const div = document.createElement("div");
    div.className = "section";
    div.innerHTML = `
      <label>Heading: <input type="text" class="heading" required></label>
      <label>Content: <textarea class="content" required></textarea></label>
    `;
    document.getElementById("sections").appendChild(div);
  }
  