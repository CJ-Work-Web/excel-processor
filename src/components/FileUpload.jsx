import React from 'react';

// I'll use a simple hidden file input with a styled label for a custom look.

import { Upload, FileSpreadsheet, CheckCircle, AlertCircle } from 'lucide-react';
import clsx from 'clsx';

export default function FileUpload({ label, file, onFileSelect, accept = ".xlsx, .xls" }) {
    const handleFileChange = (e) => {
        if (e.target.files && e.target.files[0]) {
            onFileSelect(e.target.files[0]);
        }
    };

    return (
        <div className="w-full">
            <label className="block text-sm font-medium text-gray-700 mb-2">
                {label}
            </label>

            <div className={clsx(
                "relative rounded-lg border-2 border-dashed transition-all duration-200 ease-in-out p-6 flex flex-col items-center justify-center text-center cursor-pointer hover:bg-gray-50",
                file ? "border-green-400 bg-green-50" : "border-gray-300"
            )}>
                <input
                    type="file"
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                    accept={accept}
                    onChange={handleFileChange}
                />

                {file ? (
                    <>
                        <CheckCircle className="w-10 h-10 text-green-500 mb-3" />
                        <p className="text-sm font-medium text-gray-900">{file.name}</p>
                        <p className="text-xs text-gray-500 mt-1">{(file.size / 1024).toFixed(1)} KB</p>
                    </>
                ) : (
                    <>
                        <FileSpreadsheet className="w-10 h-10 text-gray-400 mb-3" />
                        <p className="text-sm font-medium text-gray-900">點擊或拖曳檔案至此</p>
                        <p className="text-xs text-gray-500 mt-1">支援 Excel (.xlsx, .xls)</p>
                    </>
                )}
            </div>
        </div>
    );
}
