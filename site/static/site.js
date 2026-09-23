(function () {
  "use strict";

  var themeKey = "spfx-data-kit-theme";
  var validThemes = ["light", "auto", "dark"];

  function saveTheme(theme) {
    try { localStorage.setItem(themeKey, theme); } catch (_) {}
  }

  function setTheme(theme) {
    if (validThemes.indexOf(theme) === -1) theme = "auto";
    document.documentElement.dataset.theme = theme;
    saveTheme(theme);
    document.querySelectorAll('input[name="theme"]').forEach(function (input) {
      input.checked = input.value === theme;
    });
  }

  var savedTheme = "auto";
  try {
    if (validThemes.indexOf(localStorage.getItem(themeKey)) !== -1) savedTheme = localStorage.getItem(themeKey);
  } catch (_) {}
  setTheme(savedTheme);

  document.querySelectorAll('input[name="theme"]').forEach(function (input) {
    input.addEventListener("change", function () { setTheme(input.value); });
  });

  function legacyCopy(text) {
    var area = document.createElement("textarea");
    area.value = text;
    area.setAttribute("readonly", "");
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.appendChild(area);
    area.select();
    var copied = false;
    try { copied = document.execCommand("copy"); } catch (_) { copied = false; }
    document.body.removeChild(area);
    return copied;
  }

  document.querySelectorAll("[data-copy-target]").forEach(function (button) {
    button.addEventListener("click", function () {
      var source = document.getElementById(button.getAttribute("data-copy-target"));
      var status = button.parentElement.querySelector(".copy-status");
      var label = button.querySelector(".copy-label");
      var text = source ? source.innerText.replace(/\n$/, "") : "";
      var done = function (copied) {
        status.textContent = copied ? "Copied" : "Select text";
        label.textContent = copied ? "Copied" : "Copy";
        window.setTimeout(function () { status.textContent = ""; label.textContent = "Copy"; }, 2400);
      };

      if (!text) done(false);
      else if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(text).then(function () { done(true); }, function () { done(legacyCopy(text)); });
      else done(legacyCopy(text));
    });
  });
}());
