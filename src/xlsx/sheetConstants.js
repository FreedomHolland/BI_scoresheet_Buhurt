const TITLE_ROW_BACKGROUND_COLOR_ARGB = 'ffabcb95'
const YELLOW_HIGHLIGHT = 'FFFF0000'
const STANDARD_CELL_OUTLINE = '0F0F0F'
const ARGB_GREEN = 'FF00FF00'
const ARGB_RED = 'FFFF0000'
const ARGB_WHITE = 'FFFFFFFF'
const ARGB_LIGHT_GRAY = 'FFD9DADB'
const ARGB_YELLOW = 'FFFF00'


// the first row on the results sheet that we show data:
const RESULTS_TEAM_ROW_START = 7
// the first row on the pools sheet that we show data:
const POOLS_TEAM_ROW_START = 3
// the first row on the brackets sheet that we show data:
const BRACKETS_TEAM_ROW_START = 3

// how many 'fight rows' to pre-populate:
const NUMBER_FIGHTS_POOLS = 120

const TOURNAMENT_TYPE_TEAMS = 'teams'
const TOURNAMENT_TYPE_FIGHTERS = 'fighters'

const COL_SIZE = {
  xxs: 3.5,
  xs: 4.5,
  s: 6,
  m: 10,
  ml: 12,
  l: 15,
  xl: 21
}

const RESULTS_COLS = [
  { key: 'A', width: COL_SIZE.s },
  { key: 'B', width: COL_SIZE.xl },
  { key: 'C', width: COL_SIZE.s },
  { key: 'D', width: COL_SIZE.m },
  { key: 'E', width: COL_SIZE.m },
  { key: 'F', width: COL_SIZE.m },
  { key: 'G', width: COL_SIZE.m },
  { key: 'H', width: COL_SIZE.m },
  { key: 'I', width: COL_SIZE.m },
  { key: 'J', width: COL_SIZE.m },
  { key: 'K', width: COL_SIZE.m },
  { key: 'L', width: COL_SIZE.m },
  { key: 'M', width: COL_SIZE.m },
  { key: 'N', width: COL_SIZE.m },
  { key: 'O', width: COL_SIZE.m },
  { key: 'P', width: COL_SIZE.m },
  { key: 'Q', width: COL_SIZE.m },
  { key: 'R', width: COL_SIZE.m },
  { key: 'S', width: COL_SIZE.m },
  { key: 'T', width: COL_SIZE.s },
  { key: 'U', width: COL_SIZE.s },
  { key: 'V', width: COL_SIZE.s },
  { key: 'W', width: COL_SIZE.m },
  { key: 'X', width: COL_SIZE.ml },
  { key: 'Y', width: COL_SIZE.m },
  { key: 'Z', width: COL_SIZE.l },
  { key: 'AA', width: COL_SIZE.l }
]

const POOLS_BRACKETS_COLS = [
  { key: 'A', width: COL_SIZE.m },
  { key: 'B', width: COL_SIZE.s },
  { key: 'C', width: COL_SIZE.xl },
  { key: 'D', width: COL_SIZE.xs },
  { key: 'E', width: COL_SIZE.xs },
  { key: 'F', width: COL_SIZE.xs },
  { key: 'G', width: COL_SIZE.xs },
  { key: 'H', width: COL_SIZE.xs },
  { key: 'I', width: COL_SIZE.xs },
  { key: 'J', width: COL_SIZE.xs },
  { key: 'K', width: COL_SIZE.m },
  { key: 'L', width: COL_SIZE.m },
  { key: 'M', width: COL_SIZE.m },
  { key: 'N', width: COL_SIZE.s },
  { key: 'O', width: COL_SIZE.s },
  { key: 'P', width: COL_SIZE.s },
  { key: 'Q', width: COL_SIZE.s },
  { key: 'R', width: COL_SIZE.s },
  { key: 'S', width: COL_SIZE.s },
  {Key: 'T', width: COL_SIZE.xl}
]

const POOL_BRACKET_HEADER_VALUES = {
  A: 'Match#',
  B: 'ID#',
  C: 'Team',
  D: 'Rounds Score',
  I: 'Fight',
  K: 'Rounds',
  O: 'Active/Grounded',
  R: 'Cards'
}

const POOL_BRACKET_SUB_HEADER_VALUES = {
  D: 'R1',
  E: 'R2',
  F: 'R3',
  G: 'R4',
  H: 'R5',
  I: 'Won',
  J: 'Lost',
  K: 'Won',
  L: 'Draw',
  M: 'Lost',
  N: 'Ratio',
  O: 'A',
  P: 'G',
  Q: 'Ratio',
  R: '',
  S: '',
  T: 'Description',
}

const getFont = (size, bold) => {
  return {
    name: 'Cambria',
    color: { argb: '00000000' },
    family: 2,
    size: size ?? 11,
    italic: false,
    bold: !!bold,
    alignment: { horizontal: 'center', vertical: 'middle' }
  };
}

const getFontExtended = (size = 11, bold = false, italic = false, color = '00000000', alignment = {}) => {
  return {
    name: 'Cambria',
    color: { argb: color },
    family: 2,
    size: size,
    italic: italic,
    bold: bold,
    alignment: {horizontal: 'center', vertical: 'middle', wrapText: true, ...alignment}
  };
};

const TITLE_ROW_PROPS = {
  fill: {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: TITLE_ROW_BACKGROUND_COLOR_ARGB }
  },
  alignment: { horizontal: 'center', vertical: 'middle' }
  }

  const TEAM_ROW_HIGHLIGHT_PROPS = {
    fill: {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: YELLOW_HIGHLIGHT }
    },
    alignment: { horizontal: 'center', vertical: 'middle' }
  }

const EVENT_TIERS = [
  'Classic', 'Regional', 'Conference'
]

const FINALS_TYPE = [
  'Round Robin', 'Bracket'
]

const TEST_DATA = {
  eventName: 'test event',
  date: '1/2/3',
  location: 'St. Paul',
  tournamentType: 'teams',
  teams: [
    { name: 'wyverns', id: 'tha best' },
    { name: 'dfc', id: 'tha' },
    { name: 'knyaz', id: 'big boiz' }
  ],
  fighters: [
    { name: 'Derek', id: 'doc' },
    { name: 'Keegan', id: 'too big' },
    { name: 'Linden', id: 'lich boi' },
    { name: 'Vish', id: '420' },
    { name: 'Spence', id: 'coach' }
  ]
}
export {
  RESULTS_COLS,
  POOLS_BRACKETS_COLS,
  TITLE_ROW_PROPS,
  getFont,
  getFontExtended,
  EVENT_TIERS,
  FINALS_TYPE,
  RESULTS_TEAM_ROW_START,
  POOLS_TEAM_ROW_START,
  BRACKETS_TEAM_ROW_START,
  NUMBER_FIGHTS_POOLS,
  TOURNAMENT_TYPE_TEAMS,
  TOURNAMENT_TYPE_FIGHTERS,
  TEST_DATA,
  POOL_BRACKET_SUB_HEADER_VALUES,
  POOL_BRACKET_HEADER_VALUES,
  TEAM_ROW_HIGHLIGHT_PROPS,
  STANDARD_CELL_OUTLINE,
  ARGB_GREEN,
  ARGB_RED,
  ARGB_WHITE,
  ARGB_LIGHT_GRAY,
  ARGB_YELLOW
}
