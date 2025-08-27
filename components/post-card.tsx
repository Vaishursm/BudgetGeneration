import type { Post } from "@/lib/api"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Badge } from "@/components/ui/badge"
import Link from "next/link"

interface PostCardProps {
  post: Post
}

export function PostCard({ post }: PostCardProps) {
  const authorName = post.first_name && post.last_name ? `${post.first_name} ${post.last_name}` : post.username

  return (
    <Card className="hover:shadow-md transition-shadow">
      <CardHeader>
        <div className="flex items-start justify-between">
          <CardTitle className="text-lg">
            <Link href={`/posts/${post.id}`} className="hover:text-primary transition-colors">
              {post.title}
            </Link>
          </CardTitle>
          <Badge variant={post.status === "published" ? "default" : "secondary"}>{post.status}</Badge>
        </div>
        <div className="text-sm text-muted-foreground">
          By {authorName} • {new Date(post.created_at).toLocaleDateString()}
        </div>
      </CardHeader>
      <CardContent>
        <p className="text-muted-foreground line-clamp-3">{post.content.substring(0, 150)}...</p>
        {post.categories && post.categories.length > 0 && (
          <div className="flex flex-wrap gap-2 mt-4">
            {post.categories.map((category) => (
              <Badge key={category.id} variant="outline" className="text-xs">
                {category.name}
              </Badge>
            ))}
          </div>
        )}
      </CardContent>
    </Card>
  )
}
