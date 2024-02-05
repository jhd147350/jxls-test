import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.SheetData;
import org.jxls.transform.poi.PoiCellData;
import org.jxls.transform.poi.PoiSheetData;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.CannotOpenWorkbookException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * As a workaround, skip findAndRemoveExistingCellRegion(),
 * and use addMergedRegionUnsafe() instead of addMergedRegion().
 */
public class PoiTransformerForQuickMerge extends PoiTransformer {

    private static final Logger logger = LoggerFactory.getLogger(PoiTransformer.class);

    public PoiTransformerForQuickMerge(Workbook workbook, boolean streaming) {
        super(workbook, streaming);
    }

    static public PoiTransformerForQuickMerge createTransformer(InputStream is, OutputStream os) {
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(is);
        } catch (Exception e) {
            throw new CannotOpenWorkbookException(e);
        }
        PoiTransformerForQuickMerge transformer = new PoiTransformerForQuickMerge(workbook, false);
        transformer.setOutputStream(os);
        transformer.setInputStream(is);
        return transformer;
    }


    @Override
    protected void transformCell(CellRef srcCellRef, CellRef targetCellRef, Context context,
                                 boolean updateRowHeightFlag, CellData cellData, Sheet destSheet, Row destRow) {
        SheetData sheetData = sheetMap.get(srcCellRef.getSheetName());
        if (!isIgnoreColumnProps()) {
            destSheet.setColumnWidth(targetCellRef.getCol(), sheetData.getColumnWidth(srcCellRef.getCol()));
        }
        if (updateRowHeightFlag && !isIgnoreRowProps()) {
            destRow.setHeight((short) sheetData.getRowData(srcCellRef.getRow()).getHeight());
        }
        Cell destCell = destRow.getCell(targetCellRef.getCol());
        if (destCell == null) {
            destCell = destRow.createCell(targetCellRef.getCol());
        }
        try {
            // conditional formatting
            destCell.setCellType(CellType.BLANK);
            ((PoiCellData) cellData).writeToCell(destCell, context, this);
            copyMergedRegionsUnsafe(cellData, targetCellRef);
        } catch (Exception e) {
            logger.error("Failed to write a cell with {} and context keys {}", cellData, context.toMap().keySet(), e);
        }
    }

    protected final void copyMergedRegionsUnsafe(CellData sourceCellData, CellRef destCell) {
        if (sourceCellData.getSheetName() == null) {
            throw new IllegalArgumentException("Sheet name is null in copyMergedRegions");
        }
        PoiSheetData sheetData = (PoiSheetData) sheetMap.get(sourceCellData.getSheetName());
        CellRangeAddress cellMergedRegion = null;
        for (CellRangeAddress mergedRegion : sheetData.getMergedRegions()) {
            if (mergedRegion.getFirstRow() == sourceCellData.getRow() && mergedRegion.getFirstColumn() == sourceCellData.getCol()) {
                cellMergedRegion = mergedRegion;
                break;
            }
        }
        if (cellMergedRegion != null) {
            //skip findAndRemoveExistingCellRegion
            //findAndRemoveExistingCellRegion(destCell);
            Sheet destSheet = this.getWorkbook().getSheet(destCell.getSheetName());
            //destSheet.addMergedRegion(new CellRangeAddress(destCell.getRow(), destCell.getRow() + cellMergedRegion.getLastRow() - cellMergedRegion.getFirstRow(),
            //        destCell.getCol(), destCell.getCol() + cellMergedRegion.getLastColumn() - cellMergedRegion.getFirstColumn()));
            //use addMergedRegionUnsafe for faster merge operation
            destSheet.addMergedRegionUnsafe(new CellRangeAddress(destCell.getRow(), destCell.getRow() + cellMergedRegion.getLastRow() - cellMergedRegion.getFirstRow(),
                    destCell.getCol(), destCell.getCol() + cellMergedRegion.getLastColumn() - cellMergedRegion.getFirstColumn()));
        }
    }
}
