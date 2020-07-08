//@ts-check

const { JSDOM } = require('jsdom')
const path = require('path')
const fs = require('fs')

const filePath = (path.join(__dirname, '/file/{EDIT HERE}.xls'))

console.log(filePath)


async function main() {
	const dom = await JSDOM.fromFile(filePath)
	const document = dom.window.document

	const tr = [...document.querySelectorAll('tbody')[0].querySelectorAll('tr')]
	const existStudent = tr.length - 2
	const students = []
	for (let index = 10; index < existStudent; index++) {
		const row = tr[index]
		// 0 เลขที่ - 1 รหัส - 2 ชื่อ - 3 เอก
		const cells = row.cells
		students.push({
			id: cells[0].textContent.trim(),
			code: cells[1].textContent.trim(),
			name: cells[2].textContent.trim(),
			major: cells[3].textContent.trim()
		})
	}
	console.log(students)
	fs.writeFileSync(path.join(__dirname, 'out.json'), JSON.stringify(students, null, 2))

}

main()