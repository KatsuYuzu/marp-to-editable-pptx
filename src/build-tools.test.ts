import fs from 'node:fs'
import path from 'node:path'

describe('native PPTX build tasks', () => {
  it('generates the DOM walker before rebuilding every combined native bundle task', () => {
    const tasksPath = path.resolve(__dirname, '..', '.vscode', 'tasks.json')
    const tasks = JSON.parse(fs.readFileSync(tasksPath, 'utf8')) as {
      tasks: { command?: string; label?: string }[]
    }
    const combinedTasks = tasks.tasks.filter(
      (task) =>
        task.command?.includes('generate-dom-walker-script.js') &&
        task.command.includes('build-native-pptx-bundle.js'),
    )

    expect(combinedTasks.length).toBeGreaterThan(0)
    for (const task of combinedTasks) {
      const generateAt = task.command!.indexOf('generate-dom-walker-script.js')
      const buildAt = task.command!.indexOf('build-native-pptx-bundle.js')
      expect({ label: task.label, generateAt, buildAt }).toEqual(
        expect.objectContaining({
          generateAt: expect.any(Number),
          buildAt: expect.any(Number),
        }),
      )
      expect(generateAt).toBeLessThan(buildAt)
    }
  })
})
