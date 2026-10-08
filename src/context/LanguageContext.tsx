import React, { createContext, useContext, useState } from 'react';
import { UiLanguage, UI_LANGUAGES, LanguageOption, getTranslation } from '../services/i18n';

interface LanguageContextType {
  language: UiLanguage;
  setLanguage: (lang: UiLanguage) => void;
  t: (key: string, fallback?: string) => string;
  availableLanguages: LanguageOption[];
}

const LanguageContext = createContext<LanguageContextType>({
  language: 'cs',
  setLanguage: () => {},
  t: (key, fallback) => fallback || key,
  availableLanguages: UI_LANGUAGES,
});

export const LanguageProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [language, setLanguageState] = useState<UiLanguage>(() => {
    try {
      const saved = localStorage.getItem('app_ui_lang');
      if (saved && UI_LANGUAGES.some((l) => l.code === saved)) {
        return saved as UiLanguage;
      }
    } catch (_) {}
    return 'cs';
  });

  const setLanguage = (lang: UiLanguage) => {
    setLanguageState(lang);
    try {
      localStorage.setItem('app_ui_lang', lang);
    } catch (_) {}
  };

  const t = (key: string, fallback?: string) => getTranslation(language, key, fallback);

  return (
    <LanguageContext.Provider value={{ language, setLanguage, t, availableLanguages: UI_LANGUAGES }}>
      {children}
    </LanguageContext.Provider>
  );
};

export const useLanguage = () => useContext(LanguageContext);
