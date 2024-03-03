import Debug from 'debug'

// Initialize debug logging module
export const log = Debug("northwinddb-msteamext");
log.log = console.log.bind(console)
