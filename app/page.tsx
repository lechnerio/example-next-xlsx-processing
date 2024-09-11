"use client"

import { Alert, AlertDescription } from "@/components/ui/alert"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { ChangeEvent, FormEvent, useState } from "react"
import * as XLSX from "xlsx"

const Home = () => {
  const [file, setFile] = useState<File | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [firstRow, setFirstRow] = useState<
    (string | number | boolean | null)[] | null
  >(null)

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0])
      setError(null)
    }
  }

  const handleSubmit = (e: FormEvent) => {
    e.preventDefault()

    if (!file) {
      setError("Please select a file")
      return
    }

    if (file) {
      const fileExtension = file.name.split(".").pop()?.toLowerCase()

      if (!["xlsx", "xls", "csv"].includes(fileExtension!)) {
        setError("Please select a valid file")
        return
      }

      console.log("Valid file selected: ", file)
    }

    const reader = new FileReader()

    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer)
      const workbook = XLSX.read(data, { type: "array" })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]

      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
      })

      const firstRowData = jsonData[0] as (string | number | boolean | null)[]
      setFirstRow(firstRowData)
    }

    reader.readAsArrayBuffer(file)
  }

  return (
    <Card className="max-w-2xl mx-auto m-5 p-5">
      <CardHeader>
        <CardTitle>Example Next XLSX Processing</CardTitle>
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
          <Button type="submit" className="mt-4">
            Upload
          </Button>
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
      </CardContent>
    </Card>
  )
}

export default Home
