export const formatDate = (date: Date, locale: string = 'en-US'): string => {
    return new Intl.DateTimeFormat(locale).format(date);
};

export const generateRandomId = (length: number = 10): string => {
    return Math.random().toString(36).substr(2, length);
};

export const debounce = (func: Function, delay: number): Function => {
    let timeoutId: NodeJS.Timeout;
    return (...args: any[]) => {
        if (timeoutId) {
            clearTimeout(timeoutId);
        }
        timeoutId = setTimeout(() => {
            func(...args);
        }, delay);
    };
};