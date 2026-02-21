export function parseNumberParam(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

export function parseBooleanParam(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback;
  if (['1', 'true', 'yes', 'on'].includes(String(value).toLowerCase())) return true;
  if (['0', 'false', 'no', 'off'].includes(String(value).toLowerCase())) return false;
  return fallback;
}
