// Настройки тренажёра. Заполняет Назар один раз (см. NAZAR.md).
window.TRAINER_CONFIG = {
  // ID агента ElevenLabs (Agents → агент → ID, начинается с "agent_")
  agentId: "agent_XXXXXXXXXXXXXXXXXXXXXXXXXX",

  // Голос мамы для каждого языка (Voice ID из ElevenLabs).
  // Если поле пустое, используется голос по умолчанию из настроек агента.
  voices: {
    ua: "",
    ro: "",
    pl: "",
    it: "",
    en: "",
    es: "",
    sk: "",
    sr: "",
    he: ""
  },

  // Откуда брать сценарии. Основной источник — веб-приложение Apps Script
  // (адрес вида https://script.google.com/macros/s/.../exec). Оно отдаёт последний
  // файл, который прототип опубликовал в папку на Google Drive.
  scenariosUrl: "https://script.google.com/macros/s/AKfycbytEih3FpaJswwMSh9jkwe28xYdE_GH5su36RnpKGCFn8nPNb-88ojXELJp9sY61PM3/exec",

  // Запасной вариант: файл scenarios.dat в этом репозитории.
  scenariosFile: "scenarios.dat",

  // Ключ, которым прототип шифрует файл сценариев. Менять только вместе с прототипом.
  scenariosKey: "allright-trainer-v1"
};
