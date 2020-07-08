//@ts-check

const { JSDOM } = require('jsdom')
const path = require('path')
const fs = require('fs')

const fileName = '/files/Course___513103-2560_GENERAL_CHEMISTRY_LABORATORY_I_Sec_1.xls'

const file_arg = process.argv[2] // .files/Course.xls
const filePath = file_arg ? path.join(file_arg) : path.join(__dirname, fileName)
console.log(filePath)


async function main() {
	const dom = await JSDOM.fromFile(filePath)
	const document = dom.window.document

	const tr = [...document.querySelectorAll('tbody')[0].querySelectorAll('tr')]
	const students = []
	for (let index = 0; index < tr.length; index++) {
		const row = tr[index]
		// 0 เลขที่ - 1 รหัส - 2 ชื่อ - 3 เอก
		const [id, code, name, major] = row.cells
		if (+(id.textContent) > 0 && code && name && major)
			students.push({
				id: id.textContent.trim(),
				code: code.textContent.trim(),
				name: name.textContent.trim(),
				major: major.textContent.trim()
			})
	}
	// console.log(students)
	fs.writeFileSync(path.join(__dirname, 'out.json'), JSON.stringify(students, null, 2))

}

main()