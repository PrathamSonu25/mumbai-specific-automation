document.addEventListener("DOMContentLoaded", function () {
  // 1. Evidence Category Toggle
  const categorySelect = document.getElementById("evidence_category");
  const mobileSection = document.getElementById("mobile-specific-section");
  const nonMobileSection = document.getElementById("non-mobile-section");
  const subEvidencesSection = document.getElementById("sub-evidences-section");
  const extraShaSection = document.getElementById("extra-sha-section");
  const colorGroup = document.getElementById("form-group-color");
  const cloudBanner = document.getElementById("cloud-info-banner");

  function updateCategoryFields() {
    if (!categorySelect) return;
    const cat = categorySelect.value;

    const isCloudOrOther = ["Other", "Cloud", "File/Folder", "FileFolder"].includes(cat);
    const isMobile = (cat === "Mobile");
    const isComputer = ["Laptop", "Desktop", "Server"].includes(cat);

    // 1. Mobile section & extra SHA hashes
    if (mobileSection) mobileSection.style.display = isMobile ? "contents" : "none";
    if (extraShaSection) extraShaSection.style.display = isMobile ? "contents" : "none";

    // 2. Boot & Shutdown State (Only for Laptop, Desktop, Server)
    if (nonMobileSection) {
      nonMobileSection.style.display = isComputer ? "contents" : "none";
    }

    // 3. Sub-evidences (Internal Drives) (Only for Laptop, Desktop, Server)
    if (subEvidencesSection) {
      subEvidencesSection.style.display = isComputer ? "block" : "none";
    }

    // 4. Color field (Hidden for Cloud, File/Folder, Other)
    if (colorGroup) {
      colorGroup.style.display = isCloudOrOther ? "none" : "block";
    }

    // 5. Cloud guidance banner
    if (cloudBanner) {
      cloudBanner.style.display = isCloudOrOther ? "block" : "none";
    }
  }

  if (categorySelect) {
    categorySelect.addEventListener("change", updateCategoryFields);
    updateCategoryFields();
  }

  // 2. Sub-Evidence Inline Rows
  const addSubBtn = document.getElementById("btn-add-sub-evidence");
  const subTableBody = document.getElementById("sub-evidence-tbody");

  if (addSubBtn && subTableBody) {
    addSubBtn.addEventListener("click", function () {
      const row = document.createElement("tr");
      row.innerHTML = `
        <td>
          <select name="sub_type[]" class="form-control">
            <option value="SSD">SSD</option>
            <option value="HDD">HDD</option>
            <option value="Other">Other</option>
          </select>
        </td>
        <td><input type="text" name="sub_make[]" class="form-control" placeholder="Make" required /></td>
        <td><input type="text" name="sub_model[]" class="form-control" placeholder="Model" /></td>
        <td><input type="text" name="sub_sn[]" class="form-control" placeholder="Serial Number" required /></td>
        <td><input type="text" name="sub_capacity[]" class="form-control" placeholder="Capacity" /></td>
        <td><input type="text" name="sub_owner[]" class="form-control" placeholder="Owner/User" /></td>
        <td>
          <button type="button" class="btn btn-sm btn-danger btn-remove-row">&times;</button>
        </td>
      `;
      subTableBody.appendChild(row);
    });
  }

  // 3. Custody Inline Rows
  const addCustodyBtn = document.getElementById("btn-add-custody");
  const custodyTableBody = document.getElementById("custody-tbody");

  if (addCustodyBtn && custodyTableBody) {
    addCustodyBtn.addEventListener("click", function () {
      const now = new Date();
      const today = now.toISOString().split("T")[0];
      const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

      const row = document.createElement("tr");
      row.innerHTML = `
        <td>
          <input type="text" name="custody_reason[]" class="form-control" value="Forensic Backup and Hashing" required />
        </td>
        <td><input type="text" name="custody_given[]" class="form-control" placeholder="Given By Name" required /></td>
        <td><input type="text" name="custody_received[]" class="form-control" placeholder="Received By Name" required /></td>
        <td><input type="date" name="custody_date[]" class="form-control" value="${today}" /></td>
        <td><input type="text" name="custody_time[]" class="form-control" value="${timeStr}" /></td>
        <td>
          <button type="button" class="btn btn-sm btn-danger btn-remove-row">&times;</button>
        </td>
      `;
      custodyTableBody.appendChild(row);
    });
  }

  // Event delegation for row removals
  document.addEventListener("click", function (e) {
    if (e.target && e.target.classList.contains("btn-remove-row")) {
      const tr = e.target.closest("tr");
      if (tr) tr.remove();
    }
  });
});
