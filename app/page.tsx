"use client"

import { Alert, AlertDescription } from "@/components/ui/alert"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog"
import { Input } from "@/components/ui/input"
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"
import ExcelJS from "exceljs"
import { ChangeEvent, FormEvent, useState } from "react"

const Home = () => {
  const [file, setFile] = useState<File | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [firstRow, setFirstRow] = useState<
    (string | number | boolean | null)[] | null
  >(null)
  const [statusMessage, setStatusMessage] = useState<string | null>(null)
  const [rowCount, setRowCount] = useState<number>(10)
  const [availableSheets, setAvailableSheets] = useState<string[]>([]) // Store available sheet names
  const [selectedSheet, setSelectedSheet] = useState<string | null>(null) // Store the selected sheet
  const [showModal, setShowModal] = useState<boolean>(false) // Modal state
  const [replaceMode, setReplaceMode] = useState<boolean>(false) // Track if replace mode is selected
  const [userDecisionMade, setUserDecisionMade] = useState<boolean>(false) // Track if user already decided

  const [workbook, setWorkbook] = useState<ExcelJS.Workbook | null>(null) // Store workbook in state
  const [worksheet, setWorksheet] = useState<ExcelJS.Worksheet | null>(null) // Store worksheet in state

  const handleFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0])
      setError(null)

      const arrayBuffer = await e.target.files[0].arrayBuffer()
      const newWorkbook = new ExcelJS.Workbook()
      await newWorkbook.xlsx.load(arrayBuffer)

      // Collect sheet names from the workbook
      const sheetNames = newWorkbook.worksheets.map((sheet) => sheet.name)
      setAvailableSheets(sheetNames)

      // Automatically select the first worksheet by default
      if (sheetNames.length > 0) {
        setSelectedSheet(sheetNames[0])
        const firstWorksheet = newWorkbook.getWorksheet(sheetNames[0]) || null
        setWorksheet(firstWorksheet) // Store the first worksheet in state
      }

      setWorkbook(newWorkbook) // Store the workbook in state
    }
  }

  const handleSheetSelection = (value: string) => {
    setSelectedSheet(value) // Store the selected sheet
    if (workbook) {
      const newWorksheet = workbook.getWorksheet(value)
      const selectedWorksheet = newWorksheet || null
      setWorksheet(selectedWorksheet) // Update the worksheet when a new sheet is selected
    }
  }

  const handleDummyDataFill = async () => {
    if (!file || !selectedSheet) {
      setError("Please select a file and worksheet")
      return
    }

    setStatusMessage("Simulating fetching data...")

    if (worksheet) {
      // Iterate over rows starting from the second row to check if they are empty
      let hasNoEmptyRows = false

      for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
        const row = worksheet.getRow(rowIndex)

        // Check if the row is empty
        const isEmpty =
          Array.isArray(row.values) &&
          row.values.every((cell) => cell === null || cell === undefined)

        if (!isEmpty) {
          hasNoEmptyRows = true
          break
        }
      }

      // If user hasn't decided yet, show the modal
      if (hasNoEmptyRows && !userDecisionMade) {
        setShowModal(true) // Show the modal if rows are not empty
        return
      } else {
        setStatusMessage("Filling rows with random data...")
        fillRandomData(workbook!, replaceMode) // Use the selected mode (replace or append)
      }
    } else {
      setError("No visible worksheet found in the uploaded file.")
    }
  }

  const fillRandomData = (workbook: ExcelJS.Workbook, replace: boolean) => {
    if (worksheet && workbook) {
      setStatusMessage("Filling rows with random data...")

      // Get the first row and determine how many columns have data
      const firstRow = worksheet.getRow(1)
      // Filter out empty values to determine the actual last column with data
      const filledColumns = Array.isArray(firstRow.values)
        ? firstRow.values.filter(
            (cell: ExcelJS.CellValue) => cell !== null && cell !== undefined,
          )
        : []
      const lastColIndex = filledColumns.length // Number of columns with actual data

      let rowIndex = replace
        ? 2 // Start replacing from the second row if replace option is selected
        : worksheet.lastRow
        ? worksheet.lastRow.number + 1
        : 2 // Append starting after the last row with data

      for (let i = 1; i <= rowCount; i++) {
        const row = worksheet.getRow(rowIndex)
        row.values = Array(lastColIndex)
          .fill(null)
          .map(() => {
            const randomValueType = Math.floor(Math.random() * 3)
            let randomValue: string | number | boolean

            if (randomValueType === 0) {
              randomValue = getRandomString(5)
            } else if (randomValueType === 1) {
              randomValue = getRandomNumber()
            } else {
              randomValue = getRandomBoolean()
            }

            return randomValue
          })

        row.commit() // Commit the row updates
        rowIndex++
      }

      setStatusMessage("Downloading...")

      // Write the modified workbook back to the file
      if (file) {
        workbook.xlsx.writeBuffer().then((buffer) => {
          const blob = new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          })
          const link = document.createElement("a")
          link.href = URL.createObjectURL(blob)
          link.download = file.name
          link.click()
          setStatusMessage("Download complete.")
        })
      } else {
        setError("File is missing. Please upload a file first.")
      }
    }
  }

  const handleRowCountChange = (value: string) => {
    setRowCount(Number(value))
  }

  const handleSubmit = async (e: FormEvent) => {
    e.preventDefault()

    if (!file || !selectedSheet) {
      setError("Please select a file and worksheet")
      return
    }

    worksheet && setStatusMessage("Processing first row...")

    if (worksheet) {
      // Get the first row
      const firstRow = worksheet.getRow(1).values as (
        | string
        | number
        | boolean
        | null
        | ExcelJS.Cell
      )[]

      const processedFirstRow = firstRow.map((cell) => {
        if (typeof cell === "object" && cell !== null) {
          const cellTyped = cell as ExcelJS.Cell
          if ("richText" in cellTyped) {
            const richTextCell = cellTyped as { richText: { text: string }[] }
            return richTextCell.richText.map((textObj) => textObj.text).join("")
          }
          return JSON.stringify(cellTyped)
        }
        return cell
      })

      setFirstRow(processedFirstRow.filter(Boolean))
    } else {
      setError("No visible worksheet found in the uploaded file.")
    }
  }

  const getRandomString = (length: number) => {
    const characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    let result = ""
    for (let i = 0; i < length; i++) {
      result += characters.charAt(Math.floor(Math.random() * characters.length))
    }
    return result
  }

  const getRandomNumber = () => Math.floor(Math.random() * 1000)
  const getRandomBoolean = () => Math.random() < 0.5

  const handleReplace = () => {
    setReplaceMode(true) // Set to replace mode
    setUserDecisionMade(true) // Mark that the user made a decision
    setShowModal(false) // Close the modal
    setStatusMessage("Replacing rows...")

    if (workbook && selectedSheet) {
      const newWorksheet = workbook.getWorksheet(selectedSheet) // Re-fetch worksheet from workbook
      const selectedWorksheet = newWorksheet || null
      setWorksheet(selectedWorksheet) // Store worksheet in state
      if (newWorksheet) {
        fillRandomData(workbook, true) // Call fillRandomData with replace mode
      } else {
        setError("Worksheet not found in the workbook.")
      }
    } else {
      setError("Workbook or worksheet not available.")
    }
  }

  const handleAppend = () => {
    setReplaceMode(false) // Set to append mode
    setUserDecisionMade(true) // Mark that the user made a decision
    setShowModal(false) // Close the modal
    setStatusMessage("Appending rows...")

    if (workbook && selectedSheet) {
      const newWorksheet = workbook.getWorksheet(selectedSheet) // Re-fetch worksheet from workbook
      const selectedWorksheet = newWorksheet || null
      setWorksheet(selectedWorksheet) // Store worksheet in state
      if (newWorksheet) {
        fillRandomData(workbook, false) // Call fillRandomData with append mode
      } else {
        setError("Worksheet not found in the workbook.")
      }
    } else {
      setError("Workbook or worksheet not available.")
    }
  }

  return (
    <Card className="max-w-2xl mx-auto m-5 p-5">
      <CardHeader>
        <CardTitle>Example Next ExcelJS Processing</CardTitle>
      </CardHeader>

      <CardContent>
        <form onSubmit={handleSubmit}>
          <Input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileChange}
          />
          {error && (
            <Alert variant="destructive">
              <AlertDescription>{error}</AlertDescription>
            </Alert>
          )}

          <div className="mt-4 flex items-center gap-4">
            <Select onValueChange={handleRowCountChange} defaultValue="10">
              <SelectTrigger className="w-[120px]">
                <SelectValue placeholder="Select rows" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="10">10</SelectItem>
                <SelectItem value="100">100</SelectItem>
                <SelectItem value="1000">1000</SelectItem>
              </SelectContent>
            </Select>

            {availableSheets.length > 0 && (
              <div>
                <Select
                  onValueChange={handleSheetSelection}
                  value={selectedSheet || ""}
                >
                  <SelectTrigger className="w-100">
                    <SelectValue placeholder="Select worksheet" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableSheets.map((sheet) => (
                      <SelectItem key={sheet} value={sheet}>
                        {sheet}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            )}
          </div>
          {statusMessage && (
            <div className="mt-4">
              <p className="text-blue-500">{statusMessage}</p>{" "}
              {/* Adjust style as needed */}
            </div>
          )}
          <div className="mt-4 flex items-center gap-4">
            <Button variant="default" onClick={handleDummyDataFill}>
              Fill Dummy-Data & Download
            </Button>
          </div>
        </form>

        {firstRow && Array.isArray(firstRow) && (
          <div className="mt-4">
            <h2 className="font-semibold text-xl">First Row Data:</h2>
            <ul>
              {firstRow.map(
                (cell: string | number | boolean | null, index: number) => (
                  <li key={index}>{cell}</li>
                ),
              )}
            </ul>
          </div>
        )}

        {/* Modal for Replace or Append */}
        <Dialog open={showModal} onOpenChange={setShowModal}>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Existing Content Found</DialogTitle>
            </DialogHeader>
            <p>
              Do you want to replace the existing content or append new rows?
            </p>

            <div className="mt-4 flex gap-4">
              <Button variant="destructive" onClick={handleReplace}>
                Replace
              </Button>
              <Button variant="default" onClick={handleAppend}>
                Append
              </Button>
            </div>
          </DialogContent>
        </Dialog>
      </CardContent>
    </Card>
  )
}

export default Home
