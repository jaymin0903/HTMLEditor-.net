using Microsoft.EntityFrameworkCore;

namespace HTMLEditor
{
    public class EditorContext : DbContext
    {
        public EditorContext(DbContextOptions<EditorContext> options) : base(options)
        {
        }

        public DbSet<FileContent> FileContent { get; set; }
    }

    // FileContent.cs
    public class FileContent
    {
        public int Id { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string FileText { get; set; } = string.Empty;
    }

}