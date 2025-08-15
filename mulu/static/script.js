
/** 简单的中英双语切换。默认中文，点击右上角按钮在中文/英文之间切换。 */
const translations = {
  zh: {
    title: "ZTE小工具",
    subtitle: "常用内部小工具目录",
    tip: "如有问题请iCenter联系 6800000262 杜昂睿"
  },
  en: {
    title: "ZTE Toolkit",
    subtitle: "Directory of common internal tools",
    tip: "If you have any questions, please contact 6800000262 Du Angrui in iCenter"
  }
};

let currentLang = "zh";
const btn = document.getElementById("langToggle");

function applyLanguage() {
  document.querySelectorAll("[data-i18n]").forEach(el => {
    const key = el.getAttribute("data-i18n");
    if (translations[currentLang][key]) {
      el.textContent = translations[currentLang][key];
    }
  });
  // 更新每个按钮文本
  document.querySelectorAll(".btn-text").forEach(el => {
    const zh = el.getAttribute("data-i18n-key-zh");
    const en = el.getAttribute("data-i18n-key-en");
    el.textContent = currentLang === "zh" ? zh : en;
  });
  btn.textContent = currentLang === "zh" ? "EN" : "中文";
}

btn.addEventListener("click", () => {
  currentLang = currentLang === "zh" ? "en" : "zh";
  applyLanguage();
});

// 初始应用一次
applyLanguage();
