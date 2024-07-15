class TocStateManager extends TocSheet {


    constructor(state) {
        super();

        if (!Array.isArray(state)) {
            throw new Error("State must be an array");
        }

        this.currentRichTextValues = this.processArray(state);

        if (!this.currentRichTextValues.every(value => {
            console.log("VALUE: ", value);
            return value.getLinkUrl
        })) {
            throw new Error("Expected rich text values");
        }

        this.key = TocStateManager.getSheetLinksKey(); // Use static key
        this.idKey = TocStateManager.getStaticIdsKey(); // Use static key
        this.properties = TocStateManager.getScriptProps();
        this.storedState = this.getStoredState();
        this.storedIds = this.getStoredIds();

        if (!this.storedState && state.length === 0) {
            console.warn("Initial storing of state array is an empty array");
        }

        this.names = this.convertRichTextValuesToNames(this.currentRichTextValues);
        this.state = this.names; // Process initial state    
        this.currentState = this.takeSnapshot(this.state); // Deep copy of state
        this.previousState = this.storedState ? this.takeSnapshot(this.storedState) : this.takeSnapshot(this.currentState);
        this.previousIds = this.storedIds ? this.storedIds : [];
        this.currentIds = this.convertRichTextValuesToIds(this.currentRichTextValues) || [];

        // Save the initial state if it's not already stored
        if (!this.storedState) {
            this.storeState(this.state);
            this.storeIds(this.currentIds);
        }
    }


    //////////////////SIMULATES STATIC VARIABLES and STATIC METHODS/////////////////////

    static getSheetLinksKey() {
        return "sheetLinksState";
    }

    static getStaticIdsKey() {
        return "sheetLinkIds";
    }

    static getScriptProps() {
        return getPropsServ().getScriptProperties();
    }



    /**
     * Deletes the saved state and ids from the script properties.
     * 
     * This static method removes the properties identified by the static keys
     * `stateKey` and `idsKey` from the script's properties, effectively clearing
     * any previously stored state and ids.
     * 
     * Usage:
     * 
     * TocStateManager.deleteSavedState();
    */

    static deleteSavedState() {
        const properties = TocStateManager.getScriptProps();
        properties.deleteProperty(TocStateManager.getSheetLinksKey()); //Use static key
        properties.deleteProperty(TocStateManager.getStaticIdsKey()); // Uset static key
    }

    /////////////////INSTANCE METHODS////////////////

    // Snapshot allows for comparison of object and nest properties
    takeSnapshot(state) {
        return JSON.parse(JSON.stringify(state)); // Deep copy
    }

    arraysAreSimilar(arr1 = this.previousState, arr2 = this.currentState) {
        if (arr1.length !== arr2.length) {
            return false; // Arrays must have the same length
        }

        // Check if every element in arr1 is included in arr2 and vice versa
        const allElementsInArr2 = arr1.every(element => arr2.includes(element));
        const allElementsInArr1 = arr2.every(element => arr1.includes(element));

        return allElementsInArr2 && allElementsInArr1;
    }

    arraysAreIdentical(arr1 = this.previousState, arr2 = this.currentState) {
        if (arr1.length !== arr2.length) {
            return false; // Arrays must have the same length
        }

        // Check if each element in arr1 is identical to the corresponding element in arr2
        // and that each element in arr1 is in the same order as each element in arr2
        return arr1.every((element, index) => element === arr2[index]);
    }

    arraysAreSimilarButNotIdentical() {
        return this.arraysAreSimilar() && !this.arraysAreIdentical();
    }

    everyArrayElementHasALink(links = this.currentRichTextValues) {
        links = this.isArrayOfArrays(links) ? links.flatMap(ele => ele) : links;
        return links.every(ele => ele.getLinkUrl);

    }

    isArrayOfArrays(array) {
        return Array.isArray(array) ? array.every(element => Array.isArray(element)) : false;
    }

    processArray(array) {
        if (this.isArrayOfArrays(array)) {
            return array.flatMap(ele => ele);
        }
        return array;
    }

    convertRichTextValuesToNames(richTextValues) {
        return richTextValues.map(value => value.getText()) || [];
    }

    convertRichTextValuesToIds(richTextValues) {
        try {
            if (this.isArrayOfArrays(richTextValues)) {
                richTextValues = richTextValues.flatMap(ele => ele);
            }
            return richTextValues.map(value => {
                const gid = this.getSheetGIDFromRichTextUrl(value);
                if (gid === null) {
                    throw new Error(`Failed to extract GID from URL: ${value}`);
                }
                return gid;
            }) || [];
        } catch (err) {
            console.error("Caught error at id conversion: ", err.stack);
            this.restoreState(); // Ensure previous state is restored
            throw err; // Re-throw the error to propagate it upwards if necessary
        }
    }

    /**
     * Determines if both the previous and current array
     * have the same length and elements
     * @returns {Boolean}
     */
    hasStateChanged() {
        if (this.storedState) {
            return !this.arraysAreSimilar() && this.everyArrayElementHasALink()
        } else {
            return false;
        }
    }

    restoreState(previousLinks = this.getPreviousRichTextValues()) {
        // Get the range to paste previous links
        const currentRange = this.getRangeByName("TOC");
        const targetSheet = currentRange.getSheet();
        const rangeToPaste = targetSheet.getRange(currentRange.getRow(), currentRange.getColumn(), this.previousIds.length);
        // Clear ranges to paste previous rich text values
        currentRange.clear();
        rangeToPaste.clear();

        // Paste previous rich text values
        rangeToPaste.setRichTextValues(previousLinks);

        // Ensure previous named range is maintained
        this.setNamedRange("TOC", rangeToPaste);
    }

    getPreviousState() {
        return this.previousState;
    }

    getPreviousIds() {
        return this.previousIds;
    }

    getPreviousRichTextValues() {
        try {
            // Get previous Ids to make rich text values
            const previousIds = this.getPreviousIds();

            // Convert Ids to rich text value sheet links
            const getPreviousRichTextValues = this.createSheetLinks(previousIds);

            // Ensure that the rich text values are in an array of arrays
            if (!this.isArrayOfArrays(getPreviousRichTextValues)) {
                return getPreviousRichTextValues.map(value => [value]);
            }
            return getPreviousRichTextValues;
        } catch (err) {
            console.error("TOC state compromised.  Previous TOC restored: ", err.stack);
            this.restoreState();
        }
    }

    getCurrentState() {
        return this.currentState;
    }

    getCurrentIds() {
        return this.currentIds;
    }

    getCurrentRichTextValues() {
        if (!this.isArrayOfArrays(this.currentRichTextValues)) {
            return this.currentRichTextValues.map(value => [value]);
        }
        return this.currentRichTextValues
    }



    updateState(newState) {
        this.previousState = this.takeSnapshot(this.currentState);
        this.currentRichTextValues = this.processArray(newState);
        const names = this.convertRichTextValuesToNames(this.currentRichTextValues);
        const ids = this.convertRichTextValuesToIds(this.currentRichTextValues);
        this.currentState = this.takeSnapshot(names);
        this.currentIds = ids;
        this.storeState(this.currentState);
        this.storeIds(this.currentIds);
    }

    storeState(state) {
        const processedState = this.processArray(state);
        this.properties.setProperty(this.key, JSON.stringify(processedState));
    }

    storeIds(ids) {
        if (!ids || ids.length === 0) {
            console.warn("No array of ids to save");
        }

        const processedIds = this.processArray(ids);
        this.properties.setProperty(this.idKey, JSON.stringify(processedIds));
    }

    getStoredState() {
        const state = this.properties.getProperty(this.key);
        return state ? JSON.parse(state) : null;
    }

    getStoredIds() {
        const ids = this.properties.getProperty(this.idKey);
        return ids ? JSON.parse(ids).map(str => Number(str)) : null;
    }

    logChanges() {
        if (this.hasStateChanged()) {
            console.log("State has changed.");
        } else {
            console.log("State has not changed.");
        }
    }

    logStates() {
        console.log("PREVIOUS: ", this.previousState, "CURRENT: ", this.currentState);
    }

}
