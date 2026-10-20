import re
import unicodedata

# Mapping of Arabic Indic digits to Latin digits
ARABIC_INDIC_DIGITS = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')

# Common Arabic Names Dictionary for high-accuracy translation/transliteration
COMMON_ARABIC_NAMES = {
    'محمد': 'Mohammed',
    'محمود': 'Mahmoud',
    'احمد': 'Ahmed',
    'أحمد': 'Ahmed',
    'علي': 'Ali',
    'خالد': 'Khalid',
    'عبدالله': 'Abdullah',
    'عبد الله': 'Abdullah',
    'عبدالرحمن': 'Abdulrahman',
    'عبد الرحمن': 'Abdulrahman',
    'عبدالعزيز': 'Abdulaziz',
    'عبد العزيز': 'Abdulaziz',
    'عمر': 'Omar',
    'عثمان': 'Othman',
    'يوسف': 'Yousef',
    'ابراهيم': 'Ibrahim',
    'إبراهيم': 'Ibrahim',
    'سعد': 'Saad',
    'سعود': 'Saud',
    'سلطان': 'Sultan',
    'فهد': 'Fahad',
    'فيصل': 'Faisal',
    'نايف': 'Naif',
    'بندر': 'Bandar',
    'ماجد': 'Majed',
    'تركي': 'Turki',
    'سلمان': 'Salman',
    'منصور': 'Mansour',
    'ناصر': 'Nasser',
    'صالح': 'Saleh',
    'طارق': 'Tariq',
    'هشام': 'Hisham',
    'وليد': 'Waleed',
    'ياسر': 'Yasser',
    'حسن': 'Hassan',
    'حسين': 'Hussein',
    'زياد': 'Ziyad',
    'سامي': 'Sami',
    'راشد': 'Rashid',
    'جمال': 'Jamal',
    'فاطمة': 'Fatima',
    'سارة': 'Sarah',
    'مريم': 'Maryam',
    'نورة': 'Noura',
    'ريم': 'Reem',
    'منى': 'Mona',
    'هدى': 'Huda',
    'أمل': 'Amal',
    'امل': 'Amal',
    'هند': 'Hind',
    'عبير': 'Abeer',
    'روان': 'Rawan',
    'شهد': 'Shahad',
    'لولوة': 'Lulwah',
    'العنود': 'Alanoud',
    'دلال': 'Dalal',
    'منى': 'Mona',
    'زينب': 'Zainab',
    'اسماء': 'Asma',
    'أسماء': 'Asma',
    'مها': 'Maha',
    'خلود': 'Kholoud',
}

# Phonetic transliteration mapping for fallback
ARABIC_CHAR_MAP = {
    'ا': 'a', 'أ': 'a', 'إ': 'i', 'آ': 'aa', 'ء': "'", 'ؤ': 'w', 'ئ': 'y',
    'ب': 'b', 'ت': 't', 'ث': 'th', 'ج': 'j', 'ح': 'h', 'خ': 'kh',
    'د': 'd', 'ذ': 'dh', 'ر': 'r', 'ز': 'z', 'س': 's', 'ش': 'sh',
    'ص': 's', 'ض': 'dh', 'ط': 't', 'ظ': 'z', 'ع': 'a', 'غ': 'gh',
    'ف': 'f', 'ق': 'q', 'ك': 'k', 'ل': 'l', 'م': 'm', 'ن': 'n',
    'ه': 'h', 'و': 'w', 'ي': 'y', 'ى': 'a', 'ة': 'h'
}

def normalize_digits(text: str) -> str:
    """Converts Eastern Arabic digits (٠-٩) to Latin (0-9)."""
    if not isinstance(text, str):
        text = str(text) if text is not None else ""
    return text.translate(ARABIC_INDIC_DIGITS)

def normalize_arabic(text: str) -> str:
    """
    Standardizes Arabic letters for uniform searching:
    - Removes Tashkeel (diacritics: Fatha, Damma, Kasra, Shadda, Sukun)
    - Normalizes Hamzas (أ, إ, آ -> ا)
    - Normalizes Yeh / Alef Maksura (ى -> ي)
    - Normalizes Teh Marbuta (ة -> ه)
    - Normalizes digits
    """
    if not isinstance(text, str):
        text = str(text) if text is not None else ""
    
    text = normalize_digits(text)
    # Remove Tashkeel
    tashkeel_re = re.compile(r'[\u064B-\u0652\u0670\u0640]')
    text = tashkeel_re.sub('', text)
    
    # Normalize Alefs
    text = re.sub(r'[إأآ]', 'ا', text)
    # Normalize Yeh
    text = re.sub(r'ى', 'ي', text)
    # Normalize Teh Marbuta
    text = re.sub(r'ة', 'ه', text)
    
    # Collapse multiple spaces
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def transliterate_arabic_to_english(arabic_text: str) -> str:
    """
    Translates or transliterates an Arabic name/phrase to English.
    Checks dictionary first for each word, then falls back to phonetic mapping.
    """
    if not isinstance(arabic_text, str) or not arabic_text.strip():
        return ""
    
    words = arabic_text.strip().split()
    translated_words = []
    
    for word in words:
        clean_word = normalize_arabic(word)
        # Check dictionary
        if clean_word in COMMON_ARABIC_NAMES:
            translated_words.append(COMMON_ARABIC_NAMES[clean_word])
        elif word in COMMON_ARABIC_NAMES:
            translated_words.append(COMMON_ARABIC_NAMES[word])
        else:
            # Phonetic conversion
            trans_chars = []
            for char in clean_word:
                trans_chars.append(ARABIC_CHAR_MAP.get(char, char))
            converted = "".join(trans_chars).capitalize()
            translated_words.append(converted if converted else word)
            
    return " ".join(translated_words)

def contains_arabic(text: str) -> bool:
    """Checks whether the string contains Arabic characters."""
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u0600-\u06FF]', text))

def clean_alphanumeric(text: str) -> str:
    """Extracts only alphanumeric digits and characters (ignoring spaces, dashes, dots)."""
    if not isinstance(text, str):
        text = str(text) if text is not None else ""
    text = normalize_digits(text)
    return re.sub(r'[^a-zA-Z0-9]', '', text).upper()

def partial_or_full_match(target: str, query: str, full_match_only: bool = False) -> bool:
    """
    Performs normalized matching supporting:
    - Text search
    - Partial or full number matching
    - Arabic normalization
    """
    if not query or not target:
        return False
    
    # Check cleaned alphanumeric match (great for IBAN, Account numbers, contract numbers)
    target_clean = clean_alphanumeric(target)
    query_clean = clean_alphanumeric(query)
    
    if query_clean and len(query_clean) >= 3:
        if full_match_only:
            if query_clean == target_clean:
                return True
        else:
            if query_clean in target_clean:
                return True
    
    # Standard text normalized match
    norm_target = normalize_arabic(target).lower()
    norm_query = normalize_arabic(query).lower()
    
    if full_match_only:
        return norm_query == norm_target
    else:
        return norm_query in norm_target
