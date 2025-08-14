# Full-Stack Starter Template

A complete starter template featuring React.js frontend, Node.js backend, and SQLite database. Perfect for building modern web applications with authentication, CRUD operations, and a beautiful UI.

## Features

- **Frontend**: React.js with TypeScript, Next.js 14, Tailwind CSS, shadcn/ui
- **Backend**: Node.js with Next.js API routes, JWT authentication
- **Database**: SQLite with better-sqlite3, proper schema design
- **Authentication**: Complete user registration and login system
- **UI Components**: Modern, accessible components with dark mode support
- **CRUD Operations**: Full post management system with categories
- **Responsive Design**: Mobile-first approach with clean, professional styling

## Quick Start

1. **Install dependencies**
   \`\`\`bash
   npm install
   \`\`\`

2. **Set up the database**
   \`\`\`bash
   npm run db:setup
   \`\`\`

3. **Start the development server**
   \`\`\`bash
   npm run dev
   \`\`\`

4. **Open your browser**
   Navigate to [http://localhost:3000](http://localhost:3000)

## Project Structure

\`\`\`
├── app/                    # Next.js app directory
│   ├── api/               # API routes (backend)
│   ├── dashboard/         # Dashboard page
│   ├── login/            # Login page
│   ├── register/         # Registration page
│   ├── posts/            # Post-related pages
│   └── layout.tsx        # Root layout
├── components/            # Reusable React components
├── contexts/             # React contexts (auth, etc.)
├── lib/                  # Utility functions and API client
├── scripts/              # Database setup scripts
└── database.sqlite       # SQLite database (created after setup)
\`\`\`

## API Endpoints

### Authentication
- `POST /api/auth/login` - User login
- `POST /api/auth/register` - User registration

### Posts
- `GET /api/posts` - Get all posts (with optional status filter)
- `POST /api/posts` - Create new post (authenticated)
- `GET /api/posts/[id]` - Get single post
- `PUT /api/posts/[id]` - Update post (authenticated, author only)
- `DELETE /api/posts/[id]` - Delete post (authenticated, author only)

### Categories
- `GET /api/categories` - Get all categories

## Database Schema

The SQLite database includes the following tables:
- **users** - User accounts with authentication
- **posts** - Blog posts with status and author relationship
- **categories** - Post categories
- **post_categories** - Many-to-many relationship between posts and categories

## Environment Variables

Create a `.env.local` file for production:

\`\`\`env
JWT_SECRET=your-super-secret-jwt-key-change-this-in-production
\`\`\`

## Sample Data

The database setup script includes sample data:
- 3 sample users (admin, johndoe, janesmith)
- 4 sample posts with different statuses
- 4 categories (Technology, Lifestyle, Business, Education)

**Default login credentials:**
- Email: `admin@example.com`
- Password: `password123`

## Customization

### Styling
- Modify `app/globals.css` for global styles
- Update color scheme in CSS custom properties
- Components use Tailwind CSS classes

### Database
- Add new tables in `scripts/01-create-tables.sql`
- Update seed data in `scripts/02-seed-data.sql`
- Extend API operations in `lib/database.ts`

### Authentication
- JWT configuration in `lib/auth.ts`
- Add OAuth providers or other auth methods
- Customize user fields and validation

## Deployment

### Vercel (Recommended)
1. Push to GitHub
2. Connect to Vercel
3. Set environment variables
4. Deploy

### Other Platforms
1. Build the project: `npm run build`
2. Upload the SQLite database
3. Set environment variables
4. Start with: `npm start`

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

MIT License - feel free to use this template for your projects!

## Support

If you encounter any issues or have questions:
1. Check the GitHub issues
2. Create a new issue with detailed information
3. Include error messages and steps to reproduce

---

Built with ❤️ using React.js, Node.js, and SQLite
