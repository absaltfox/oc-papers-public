import path from 'node:path';

export const DATA_DIR = process.env.APP_DATA_DIR || path.join(process.cwd(), 'data');
export const CACHE_TTL_MS = Number(process.env.CACHE_TTL_MS || 10 * 60 * 1000);

export const STOP_WORDS = new Set([
  'about', 'after', 'again', 'against', 'among', 'also', 'been', 'before', 'being', 'between',
  'both', 'can', 'could', 'did', 'does', 'doing', 'during', 'each', 'from', 'have', 'having',
  'here', 'into', 'itself', 'just', 'more', 'most', 'much', 'must', 'only', 'other', 'over',
  'same', 'should', 'some', 'such', 'than', 'that', 'their', 'theirs', 'them', 'then', 'there',
  'these', 'they', 'this', 'those', 'through', 'under', 'until', 'very', 'were', 'what', 'when',
  'where', 'which', 'while', 'with', 'within', 'without', 'would', 'your', 'yours', 'study',
  'research', 'thesis', 'dissertation', 'ubc', 'university', 'doctoral', 'doctor', 'education'
]);
