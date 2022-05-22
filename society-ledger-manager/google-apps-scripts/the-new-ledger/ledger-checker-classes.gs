/*
  =============================================================================
  Project Page: https://github.com/cmenon12/contemporary-choir
  Copyright:    (c) 2022 by Christopher Menon
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  =============================================================================
 */


/**
 * Represents a single entry on a ledger.
 */
class Entry {

  /**
   * Constructor, sets everything needed.
   *
   * @param {String} date
   * @param {String} description
   * @param {Number} money
   */
  constructor(date, description, money) {
    Object.assign(this, {date, description, money});
  }

}


/**
 * Represents a cost code on a ledger, including an array of entries.
 */
class CostCode {

  /**
   * Constructor, sets everything needed.
   * Checks that balance===moneyIn-moneyOut.
   *
   * @param {Number} moneyIn
   * @param {Number} moneyOut
   * @param {Number} balance
   * @param {Number} lastRowNumber
   */
  constructor(moneyIn, moneyOut, balance, lastRowNumber) {
    if (balance !== moneyIn - moneyOut) {
      Logger.log(`balance=${balance}; moneyIn=${moneyIn}; moneyOut=${moneyOut}`);
      // throw new Error("The balance is invalid (balance!==moneyIn-moneyOut).");
    }
    const entries = [];
    Object.assign(this, {moneyIn, moneyOut, balance, lastRowNumber, entries});
    this.changeInBalance = 0;
  }

  /**
   * Adds an entry to the cost code, and increments the change in
   * the balance for the cost code.
   *
   * @param {Entry} entry the entry to add.
   */
  addEntry(entry) {
    this.entries.push(entry);
    this.changeInBalance = this.changeInBalance + entry.money;

  }

}


/**
 * Represents the grand total for a cost code.
 */
class GrandTotal {

  /**
   * Constructor, sets everything needed.
   * Checks that balance===balanceBroughtForward+totalIn-totalOut.
   *
   * @param {Number} totalIn
   * @param {Number} totalOut
   * @param {Number} balanceBroughtForward
   * @param {Number} totalBalance
   * @param {Number} lastRowNumber
   */
  constructor(totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber) {
    if (totalBalance !== balanceBroughtForward + totalIn - totalOut) {
      Logger.log(`totalBalance=${totalBalance}; balanceBroughtForward=${balanceBroughtForward}; totalIn=${totalIn}; totalOut=${totalOut}`);
      // throw new Error("The total balance is invalid (totalBalance!==balanceBroughtForward+totalIn-totalOut).");
    }
    Object.assign(this, {totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber});
  }

}


/**
 * Represents a ledger.
 * It contains the sheet ID, society name, an array of cost codes,
 * and the grand total.
 * The entries are contained within each cost code.
 */
class Ledger {

  /**
   * Constructor, sets the sheetID, creates empty arrays for entries
   * and costCodes, and sets grandTotal as undefined.
   *
   * @param {Number} sheetId the sheet ID that this represents
   */
  constructor(sheetId) {
    this.sheetId = sheetId;
    this.costCodes = {};
    this.grandTotal = undefined;
  }

  /**
   * Calculates a single value for the money based on the money in
   * or money out (which is how the ledger represents it).
   *
   * @param {Number} moneyIn the money in
   * @param {Number} moneyOut the money out
   * @return {Number} the actual change
   */
  static calculateMoney(moneyIn, moneyOut) {

    if (Number(moneyIn) === 0) {
      return -moneyOut;
    } else {
      return moneyIn;
    }

  }

  /**
   * Compares the two ledgers.
   * Returns true if their grand totals have the same totalIn,
   * totalOut, and balanceBroughtForward.
   *
   * @param {Ledger} ledgerA one ledger.
   * @param {Ledger} ledgerB another ledger.
   * @return {boolean} true if the same, otherwise false.
   */
  static compareLedgers(ledgerA, ledgerB) {

    return ledgerA.grandTotal.totalIn === ledgerB.grandTotal.totalIn &&
      ledgerA.grandTotal.totalOut === ledgerB.grandTotal.totalOut &&
      ledgerA.grandTotal.balanceBroughtForward === ledgerB.grandTotal.balanceBroughtForward;

  }

  /**
   * Adds a new entry to the named cost code on this ledger.
   *
   * @param {String} costCodeName
   * @param {String} date
   * @param {String} description
   * @param {Number} moneyIn
   * @param {Number} moneyOut
   */
  addEntry(costCodeName, date, description, moneyIn = undefined, moneyOut = undefined) {
    const money = Ledger.calculateMoney(moneyIn, moneyOut);

    if (!this.costCodes.hasOwnProperty(costCodeName)) {
      throw new Error(`The cost code ${costCodeName} doesn't exist.`);
    } else {
      this.costCodes[costCodeName].addEntry(new Entry(date, description, money));
    }

  }

  /**
   * Adds a new cost code.
   *
   * @param {String} name
   * @param {Number} moneyIn
   * @param {Number} moneyOut
   * @param {Number} balance
   * @param {Number} lastRowNumber
   */
  addCostCode(name, moneyIn, moneyOut, balance, lastRowNumber) {
    Object.assign(this.costCodes,
      {[name]: new CostCode(moneyIn, moneyOut, balance, lastRowNumber)});
  }

  /**
   * Sets the name of the society.
   *
   * @param {String} name
   */
  setSocietyName(name) {
    this.societyName = name;
  }

  /**
   * Sets the grand total. There can only be one grand total.
   *
   * @param {Number} totalIn
   * @param {Number} totalOut
   * @param {Number} balanceBroughtForward
   * @param {Number} totalBalance
   * @param {Number} lastRowNumber
   */
  setGrandTotal(totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber) {
    this.grandTotal = new GrandTotal(totalIn, totalOut, balanceBroughtForward, totalBalance, lastRowNumber);
  }

  /**
   * Sets the timestamp of the old ledger we compared against.
   *
   * @param {String} timestamp
   */
  setOldLedgerTimestamp(timestamp) {
    this.oldLedgerTimestamp = timestamp;
  }

  /**
   * Returns each cost code with its last row number.
   * It's returned as an array of cost codes, with each element as
   * [name, lastRowNumber].
   *
   * @return {[{String}|{Number}]} the array of cost codes.
   */
  getCostCodeRows() {

    const costCodeRows = [];
    for (const [key, value] of Object.entries(this.costCodes)) {
      costCodeRows.push([key, value.lastRowNumber]);
    }
    Logger.log(`costCodeRows is: ${costCodeRows}.`);
    return costCodeRows;

  }

  /**
   * Saves the whole thing to the log.
   */
  log() {
    Logger.log(`The Ledger object is: ${JSON.stringify(this)}`);
  }
}
